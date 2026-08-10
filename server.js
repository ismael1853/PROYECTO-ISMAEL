const express = require('express');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const QRCode = require('qrcode');
const fs = require('fs');
const app = express();

const path = require('path');

// Servir los archivos de la interfaz desde la carpeta public
app.use(express.static(path.join(__dirname, 'public')));

// Ruta principal para enviar el archivo index.html de CORPOELEC
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.use(express.json());

// --- MOCK DATABASE (Ejemplo de la estructura que irá en la nube) ---
let bitacoraAccesos = [];
let inventarioEquipos = [
    { 
        id: "EQ-001", tipo: "Laptop", marca: "Lenovo", modelo: "ThinkPad E14", serial: "L3X9921A",
        ram: "16GB", disco: "512GB SSD", mac: "00:1A:2B:3C:4D:5E", ip: "10.40.12.55",
        ubicacion: "Sede Centro", piso: "Piso 4",
        asignadoA: { nombre: "Pedro", apellido: "Pérez", cedula: "V-15.342.111", nPersonal: "CORP-9942", depto: "Distribución", correo: "pperez@corpoelec.gob.ve" }
    }
];

// --- 🔐 REGISTRO DE ACCESOS (BITÁCORA) ---
app.post('/api/auth/login', (req, res) => {
    const { usuario, clave } = req.body;
    
    // Validación lógica de roles (Simulación)
    const acceso = {
        usuario: usuario,
        fecha: new Date().toISOString(),
        ip_origen: req.ip,
        evento: "Inicio de sesión exitoso"
    };
    
    bitacoraAccesos.push(acceso);
    console.log("Bitácora actualizada:", acceso);
    res.json({ success: true, token: "session-token-valid-corpoelec", rol: "Admin" });
});

// --- 📊 EXCEL: REPORTE DE INVENTARIO ---
app.get('/api/reportes/excel', async (req, res) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Inventario Tecnológico');

    worksheet.columns = [
        { header: 'ID', key: 'id', width: 10 },
        { header: 'Tipo', key: 'tipo', width: 15 },
        { header: 'Marca', key: 'marca', width: 15 },
        { header: 'Serial', key: 'serial', width: 20 },
        { header: 'Ubicación', key: 'ubicacion', width: 15 },
        { header: 'Asignado A', key: 'responsable', width: 25 }
    ];

    inventarioEquipos.forEach(eq => {
        worksheet.addRow({
            id: eq.id,
            tipo: eq.tipo,
            marca: eq.marca,
            serial: eq.serial,
            ubicacion: `${eq.ubicacion} - ${eq.piso}`,
            responsable: `${eq.asignadoA.nombre} ${eq.asignadoA.apellido} (${eq.asignadoA.cedula})`
        });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Inventario_CORPOELEC.xlsx');
    await workbook.xlsx.write(res);
    res.end();
});

// --- 📄 PDF: REPORTE CON MEMBRETE Y CÓDIGO QR ---
app.get('/api/reportes/pdf', async (req, res) => {
    const doc = new PDFDocument({ margin: 50 });
    
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename=Ficha_Inventario_Corpoelec.pdf');
    doc.pipe(res);

    // Membrete Oficial (Estructura de Texto / Logo Corporativo)
    doc.fontSize(10).text("CORPOELEC - MINISTERIO DEL PODER POPULAR PARA LA ENERGÍA ELÉCTRICA", { align: 'center' });
    doc.text("DIVISIÓN DE TELECOMUNICACIONES E INFORMÁTICA", { align: 'center' });
    doc.moveDown();
    doc.strokeColor("#5C2414").lineWidth(3).moveTo(50, 85).lineTo(550, 85).stroke(); // Línea color Marrón Corpoelec
    
    doc.moveDown(2);
    doc.fontSize(16).fillColor("#5C2414").text("FICHA DE ASIGNACIÓN DE ACTIVOS TECNOLÓGICOS", { align: 'center', bold: true });
    doc.moveDown();

    // Data del equipo
    const equipo = inventarioEquipos[0];
    doc.fontSize(11).fillColor("#000000");
    doc.text(`Tipo de Activo: ${equipo.tipo}`);
    doc.text(`Marca / Modelo: ${equipo.marca} ${equipo.modelo}`);
    doc.text(`Número de Serial: ${equipo.serial}`);
    doc.text(`Especificaciones: RAM: ${equipo.ram} | Disco: ${equipo.disco}`);
    doc.text(`Red: MAC: ${equipo.mac} | IP asignada: ${equipo.ip}`);
    doc.text(`Ubicación Física: ${equipo.ubicacion}, ${equipo.piso}`);
    
    doc.moveDown();
    doc.text("DATOS DEL USUARIO ASIGNADO:", { underline: true });
    doc.text(`Nombre y Apellido: ${equipo.asignadoA.nombre} ${equipo.asignadoA.apellido}`);
    doc.text(`Cédula / N° Personal: ${equipo.asignadoA.cedula} / ${equipo.asignadoA.nPersonal}`);
    doc.text(`Departamento / Correo: ${equipo.asignadoA.depto} - ${equipo.asignadoA.correo}`);

    // Generar e Incrustar Código QR con la metadata completa del equipo
    const qrDataText = `ID: ${equipo.id} | Serial: ${equipo.serial} | Asignado a: ${equipo.asignadoA.cedula}`;
    const qrBuffer = await QRCode.toBuffer(qrDataText);
    
    doc.image(qrBuffer, 430, 120, { width: 100 }); // Posiciona el QR arriba a la derecha
    
    doc.moveDown(4);
    doc.text("___________________________               ___________________________", { align: 'center' });
    doc.text("Entregado por (Soporte TI)                      Recibido por (Usuario)", { align: 'center' });

    doc.end();
});

app.listen(3000, () => console.log('Servidor SIV-CORPOELEC corriendo en puerto 3000'));
