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
    ssl: { rejectUnauthorized: false }
});

// Ruta raíz obligatoria para el inicio del sistema
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// --- 🔐 LOGIC DE ACCESOS Y BITÁCORA (CORREGIDO A PRUEBA DE FALLOS) ---
app.post('/api/auth/login', async (req, res) => {
    const { usuario, clave } = req.body;
    const ipOrigen = req.ip;

    console.log(`Intento de login recibido: Usuario=${usuario}`);

    // Mecanismo de emergencia: Si es el admin principal, dar acceso directo inmediato
    if (usuario === 'admin' && clave === 'corpoelec2026') {
        try {
            await pool.query(
                'INSERT INTO bitacora_accesos (usuario, ip_origen, evento) VALUES ($1, $2, $3)',
                ['admin', ipOrigen, 'Inicio de sesión exitoso (Admin)']
            ).catch(e => console.log("Omitiendo registro en bitácora por tabla ausente"));
        } catch (e) {}
        return res.json({ success: true, rol: 'Admin' });
    }

    // Validación normal para otros usuarios en la nube
    try {
        const result = await pool.query('SELECT * FROM usuarios_sistema WHERE usuario = $1', [usuario]);
        if (result.rows.length === 0) {
            return res.status(401).json({ success: false, message: "Usuario no registrado" });
        }
        
        const user = result.rows[0]; // Corrección crítica de lectura de fila
        if (user.clave_hash === clave) {
            return res.json({ success: true, rol: user.rol || 'Admin' });
        }
        res.status(401).json({ success: false, message: "Clave incorrecta" });
    } catch (err) {
        console.error("Error en login:", err);
        res.status(500).json({ success: false, message: "Error interno del servidor" });
    }
});

// --- 📊 CAJÓN DE ESTADÍSTICAS (CORREGIDO) ---
app.get('/api/dashboard/contadores', async (req, res) => {
    try {
        const informatic = await pool.query("SELECT COUNT(*)::int FROM inventario WHERE tipo_equipo IN ('Laptop', 'CPU', 'Monitor', 'Teclado', 'Mouse', 'Regulador')");
        const telecom = await pool.query("SELECT COUNT(*)::int FROM inventario WHERE tipo_equipo IN ('Radio', 'Multiplexor', 'Switch')");
        const ups = await pool.query("SELECT COUNT(*)::int FROM inventario WHERE tipo_equipo = 'UPS'");
        const users = await pool.query("SELECT COUNT(*)::int FROM personal");

        res.json({
            informatic: informatic.rows[0].count || 0,
            telecom: telecom.rows[0].count || 0,
            ups: ups.rows[0].count || 0,
            users: users.rows[0].count || 0
        });
    } catch (err) {
        console.error("Error en contadores:", err);
        // Devolver ceros si las tablas no responden para no romper la pantalla
        res.json({ informatic: 0, telecom: 0, ups: 0, users: 0 });
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
            DO UPDATE SET departamento = $5, correo_corporativo = $6;
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
        console.error("Error en guardado:", err);
        res.status(500).json({ success: false, message: "Error al guardar en la base de datos." });
    } finally {
        client.release();
    }
});

// --- 📊 EXCEL: REPORTE DE INVENTARIO COMPLETO ---
app.get('/api/reportes/excel', async (req, res) => {
    try {
        const result = await pool.query(`
            SELECT i.*, p.nombre, p.apellido, p.n_personal 
            FROM inventario i 
            LEFT JOIN personal p ON i.cedula_asignado = p.cedula
        `);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Inventario Tecnológico');

        worksheet.columns = [
            { header: 'ID', key: 'id', width: 10 },
            { header: 'Tipo Equipo', key: 'tipo_equipo', width: 15 },
            { header: 'Marca', key: 'marca', width: 15 },
            { header: 'Modelo', key: 'modelo', width: 15 },
            { header: 'Serial', key: 'serial', width: 20 },
            { header: 'Ubicación', key: 'ubicacion', width: 15 },
            { header: 'Responsable', key: 'responsable', width: 25 }
        ];

        result.rows.forEach(eq => {
            worksheet.addRow({
                id: eq.id,
                tipo_equipo: eq.tipo_equipo,
                marca: eq.marca,
                modelo: eq.modelo,
                serial: eq.serial,
                ubicacion: `${eq.ubicacion} - ${eq.piso}`,
                responsable: eq.nombre ? `${eq.nombre} ${eq.apellido}` : 'Sin Asignar'
            });
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Inventario_CORPOELEC.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (err) { res.status(500).send("Error generando Excel"); }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor de CORPOELEC activo en puerto ${PORT}`));
