require('dotenv').config();
const express = require('express');
const { Pool } = require('pg');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const QRCode = require('qrcode');
const path = require('path');

const app = express();
app.use(express.json());

// Servir la interfaz estática desde la carpeta public
app.use(express.static(path.join(__dirname, 'public')));

// Conexión a la Base de Datos en la Nube (Supabase)
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: { rejectUnauthorized: false } // Requerido por Supabase para conexiones seguras
});

// Ruta raíz obligatoria para el inicio del sistema
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// --- 🔐 LOGIC DE ACCESOS Y BITÁCORA ---
app.post('/api/auth/login', async (req, res) => {
    const { usuario, clave } = req.body;
    const ipOrigen = req.ip;
    try {
        const result = await pool.query('SELECT * FROM usuarios_sistema WHERE usuario = $1', [usuario]);
        if (result.rows.length === 0) {
            return res.status(401).json({ success: false, message: "Usuario no registrado" });
        }
        const user = result.rows[0];
        await pool.query(
            'INSERT INTO bitacora_accesos (usuario, ip_origen, evento) VALUES ($1, $2, $3)',
            [usuario, ipOrigen, `Inicio de sesión exitoso - Rol: ${user.rol}`]
        );
        res.json({ success: true, rol: user.rol });
    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, message: "Error en el servidor" });
    }
});

// --- 📊 CAJÓN DE ESTADÍSTICAS (PANTALLA PRINCIPAL) ---
app.get('/api/dashboard/contadores', async (req, res) => {
    try {
        const informatic = await pool.query("SELECT COUNT(*)::int FROM inventario WHERE tipo_equipo IN ('Laptop', 'CPU', 'Monitor', 'Teclado', 'Mouse', 'Regulador')");
        const telecom = await pool.query("SELECT COUNT(*)::int FROM inventario WHERE tipo_equipo IN ('Radio', 'Multiplexor', 'Switch')");
        const ups = await pool.query("SELECT COUNT(*)::int FROM inventario WHERE tipo_equipo = 'UPS'");
        const users = await pool.query("SELECT COUNT(*)::int FROM personal");

        res.json({
            informatic: informatic.rows[0].count,
            telecom: telecom.rows[0].count,
            ups: ups.rows[0].count,
            users: users.rows[0].count
        });
    } catch (err) {
        console.error(err);
        res.status(500).send("Error al obtener estadísticas");
    }
});

// --- 💾 ENDPOINT: REGISTRAR TRABAJADOR Y ASIGNAR EQUIPO ---
app.post('/api/inventario/registrar', async (req, res) => {
    const client = await pool.connect();
    try {
        await client.query('BEGIN');
        const {
            cedula, nombre, apellido, n_personal, departamento, correo_corporativo,
            tipo_equipo, marca, modelo, serial, ram, disco_duro, telefono_ip, mac_address, ubicacion, piso
        } = req.body;

        const queryPersonal = `
            INSERT INTO personal (cedula, nombre, apellido, n_personal, departamento, correo_corporativo)
            VALUES ($1, $2, $3, $4, $5, $6)
            ON CONFLICT (cedula) 
            DO UPDATE SET departamento = $5, correo_corporativo = $6
            RETURNING cedula;
        `;
        await client.query(queryPersonal, [cedula, nombre, apellido, n_personal, departamento, correo_corporativo]);

        const queryInventario = `
            INSERT INTO inventario (tipo_equipo, marca, modelo, serial, ram, disco_duro, telefono_ip, mac_address, ubicacion, piso, cedula_asignado)
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11);
        `;
        await client.query(queryInventario, [tipo_equipo, marca, modelo, serial, ram, disco_duro, telefono_ip, mac_address, ubicacion, piso, cedula]);

        await client.query('COMMIT');
        res.json({ success: true, message: "Registro completado con éxito" });
    } catch (err) {
        await client.query('ROLLBACK');
        console.error("Error en la transacción de guardado:", err);
        if (err.code === '23505') {
            res.status(400).json({ success: false, message: "El número de serial o ficha del personal ya se encuentra registrado." });
        } else {
            res.status(500).json({ success: false, message: "Error interno del servidor al procesar la data." });
        }
    } finally {
        client.release();
    }
});

// --- 📊 EXCEL: REPORTE DE INVENTARIO COMPLETO ---
app.get('/api/reportes/excel', async (req, res) => {
    try {
        const queryText = `
            SELECT i.*, p.nombre, p.apellido, p.n_personal, p.departamento 
            FROM inventario i 
            LEFT JOIN personal p ON i.cedula_asignado = p.cedula`;
        const result = await pool.query(queryText);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Inventario Tecnológico');

        worksheet.columns = [
            { header: 'ID', key: 'id', width: 10 },
            { header: 'Tipo Equipo', key: 'tipo_equipo', width: 15 },
            { header: 'Marca', key: 'marca', width: 15 },
            { header: 'Modelo', key: 'modelo', width: 15 },
            { header: 'Serial', key: 'serial', width: 20 },
            { header: 'Ubicación', key: 'ubicacion', width: 15 },
            { header: 'Piso', key: 'piso', width: 15 },
            { header: 'Responsable', key: 'responsable', width: 25 }
        ];

        result.rows.forEach(eq => {
            worksheet.addRow({
                id: eq.id,
                tipo_equipo: eq.tipo_equipo,
                marca: eq.marca,
                modelo: eq.modelo,
                serial: eq.serial,
                ubicacion: eq.ubicacion,
                piso: eq.piso,
                responsable: eq.nombre ? `${eq.nombre} ${eq.apellido} (${eq.n_personal})` : 'Sin Asignar'
            });
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Inventario_CORPOELEC.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        console.error(err);
        res.status(500).send("Error generando Excel");
    }
});

// --- 📄 PDF: REPORTE INDIVIDUAL CON MEMBRETE Y QR ---
app.get('/api/reportes/pdf/:id', async (req, res) => {
    const equipoId = req.params.id;
    try {
        const queryText = `
            SELECT i.*, p.nombre, p.apellido, p.n_personal, p.departamento, p.correo_corporativo 
            FROM inventario i 
            LEFT JOIN personal p ON i.cedula_asignado = p.cedula 
            WHERE i.id = $1`;
        
        const result = await pool.query(queryText, [equipoId]);
        if (result.rows.length === 0) return res.status(404).send("Equipo no encontrado");
        
        const eq = result.rows[0];
        const doc = new PDFDocument({ margin: 50 });

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename=Ficha_${eq.serial}.pdf`);
        doc.pipe(res);

        doc.fontSize(10).text("CORPOELEC - MINISTERIO DEL PODER POPULAR PARA LA ENERGÍA ELÉCTRICA", { align: 'center' });
        doc.text("DIVISIÓN DE TELECOMUNICACIONES E INFORMÁTICA", { align: 'center' });
        doc.moveDown();
        doc.strokeColor("#5C2414").lineWidth(3).moveTo(50, 85).lineTo(550, 85).stroke(); 
        
        doc.moveDown(2);
        doc.fontSize(14).fillColor("#5C2414").text("FICHA DE ASIGNACIÓN DE ACTIVOS TECNOLÓGICOS", { align: 'center', bold: true });
        doc.moveDown();

        doc.fontSize(11).fillColor("#000000");
        doc.text(`Tipo de Activo: ${eq.tipo_equipo}`);
        doc.text(`Marca / Modelo: ${eq.marca} / ${eq.modelo}`);
        doc.text(`Número de Serial: ${eq.serial}`);
        doc.text(`Especificaciones: RAM: ${eq.ram || 'N/A'} | Disco: ${eq.disco_duro || 'N/A'}`);
        doc.text(`Red: MAC: ${eq.mac_address || 'N/A'} | IP: ${eq.telefono_ip || 'N/A'}`);
        doc.text(`Ubicación Física: ${eq.ubicacion} - Piso: ${eq.piso}`);
        
        doc.moveDown();
        doc.text("DATOS DEL USUARIO RESPONSABLE:", { underline: true });
        doc.text(`Nombre y Apellido: ${eq.nombre} ${eq.apellido}`);
        doc.text(`Cédula / N° Personal: ${eq.cedula_asignado} / ${eq.n_personal}`);
        doc.text(`Departamento / Correo: ${eq.departamento} - ${eq.correo_corporativo}`);

        const qrDataText = `ID: ${eq.id} | Serial: ${eq.serial} | Responsable: ${eq.cedula_asignado}`;
        const qrBuffer = await QRCode.toBuffer(qrDataText);
        doc.image(qrBuffer, 440, 110, { width: 100 }); 
        
        doc.moveDown(4);
        doc.text("___________________________               ___________________________", { align: 'center' });
        doc.text("Entregado por (Soporte TI)                      Recibido por (Usuario)", { align: 'center' });

        doc.end();
    } catch (err) {
        console.error(err);
        res.status(500).send("Error generando PDF");
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor de CORPOELEC corriendo en puerto ${PORT}`));
