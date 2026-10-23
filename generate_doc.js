const fs = require('fs');
const docx = require('docx');

const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, ImageRun, Table, TableRow, TableCell, WidthType, BorderStyle } = docx;

async function generate() {
    const doc = new Document({
        sections: [
            {
                properties: {},
                children: [
                    // --- CARATULA ---
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                            new TextRun({ text: "UNIVERSIDAD / INSTITUCIÓN", bold: true, size: 28, color: "1B365D" })
                        ],
                        spacing: { after: 200 }
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                            new TextRun({ text: "FACULTAD DE INGENIERÍA / SEGURIDAD INDUSTRIAL", bold: true, size: 24, color: "333333" })
                        ],
                        spacing: { after: 400 }
                    }),
                    // Logo space or title
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                            new TextRun({ text: "[ ESPACIO PARA LOGO ]", italics: true, size: 20, color: "888888" })
                        ],
                        spacing: { after: 800 }
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                            new TextRun({ text: "EVALUACIÓN DE RIESGOS DE INCENDIO", bold: true, size: 36, color: "1B365D" })
                        ],
                        spacing: { after: 200 }
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                            new TextRun({ text: "Método MESERI aplicado al CEDI D1 Lebrija", bold: true, size: 28, color: "555555" })
                        ],
                        spacing: { after: 1200 }
                    }),
                    // Datos del Alumno y Profesor
                    new Paragraph({
                        alignment: AlignmentType.LEFT,
                        children: [
                            new TextRun({ text: "Profesor: ", bold: true, size: 22 }),
                            new TextRun({ text: "Ing. [Nombre del Profesor]", size: 22 })
                        ],
                        spacing: { after: 150 }
                    }),
                    new Paragraph({
                        alignment: AlignmentType.LEFT,
                        children: [
                            new TextRun({ text: "Alumno: ", bold: true, size: 22 }),
                            new TextRun({ text: "[Tu Nombre / Nombre del Alumno]", size: 22 })
                        ],
                        spacing: { after: 150 }
                    }),
                    new Paragraph({
                        alignment: AlignmentType.LEFT,
                        children: [
                            new TextRun({ text: "Asignatura: ", bold: true, size: 22 }),
                            new TextRun({ text: "Seguridad y Prevención de Incendios", size: 22 })
                        ],
                        spacing: { after: 150 }
                    }),
                    new Paragraph({
                        alignment: AlignmentType.LEFT,
                        children: [
                            new TextRun({ text: "Fecha: ", bold: true, size: 22 }),
                            new TextRun({ text: "Octubre 2026", size: 22 })
                        ],
                        pageBreakBefore: true
                    }),

                    // --- CONTENIDO DEL REPORTE ---
                    new Paragraph({
                        text: "1. Introducción y Objetivos",
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 300, after: 150 }
                    }),
                    new Paragraph({
                        text: "El presente documento detalla la evaluación de riesgos de incendio realizada en el Centro de Distribución (CEDI) D1 ubicado en Lebrija. El objetivo principal es identificar las vulnerabilidades, analizar los factores de protección y aplicar el Método MESERI (Método Estadístico de Riesgo de Incendio) para determinar el nivel de riesgo global de las instalaciones.",
                        spacing: { after: 200 }
                    }),

                    new Paragraph({
                        text: "2. Descripción de las Instalaciones (CEDI D1 Lebrija)",
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 300, after: 150 }
                    }),
                    new Paragraph({
                        text: "El CEDI D1 Lebrija es una infraestructura logística destinada al almacenamiento y distribución de mercancía masiva. Cuenta con zonas de estanterías de gran altura (racks), áreas de recepción y despacho, oficinas administrativas y zonas técnicas de soporte.",
                        spacing: { after: 200 }
                    }),

                    new Paragraph({
                        text: "3. Aplicación del Método MESERI",
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 300, after: 150 }
                    }),
                    new Paragraph({
                        text: "El Método MESERI evalúa diferentes factores de riesgo y protección. A continuación se presentan los aspectos evaluados:",
                        spacing: { after: 200 }
                    }),
                    new Paragraph({
                        text: "• Constructivos: Características de los materiales, resistencia al fuego de estructuras y sectorización.\n• Procesos: Peligrosidad de la manipulación de mercancías y cargas térmicas.\n• Protección: Disponibilidad de extornadores, redes contra incendio, rociadores automáticos y sistemas de detección temprana.\n• Factores agravantes: Altura de almacenamiento, superficie de los sectores de incendio y accesibilidad para cuerpos de bomberos.",
                        spacing: { after: 300 }
                    }),

                    new Paragraph({
                        text: "4. Evidencia Fotográfica e Inspección en Terreno",
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 300, after: 150 }
                    }),
                    new Paragraph({
                        text: "Durante la visita de inspección al CEDI D1 Lebrija se tomaron registros fotográficos del estado actual de los sistemas de seguridad, pasillos, tableros eléctricos y zonas de almacenamiento.",
                        spacing: { after: 200 }
                    }),

                    // Insert images if they exist
                    ...getImagesParagraphs(),

                    new Paragraph({
                        text: "5. Conclusiones y Recomendaciones",
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 300, after: 150 }
                    }),
                    new Paragraph({
                        text: "1. Se recomienda el mantenimiento preventivo trimestral de la red contra incendios.\n2. Mantener despejados los pasillos de evacuación y salidas de emergencia identificadas en la inspección.\n3. Capacitar de forma continua al personal en el uso correcto de extintores portátiles y protocolos de evacuación.",
                        spacing: { after: 200 }
                    })
                ]
            }
        ]
    });

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync("Evaluacion_Riesgos_Incendio_CEDI_D1_Lebrija.docx", buffer);
    console.log("Word document created successfully!");
}

function getImagesParagraphs() {
    const paras = [];
    const fsObj = require('fs');
    if (fsObj.existsSync('imagenes_cdi_reporte')) {
        const files = fsObj.readdirSync('imagenes_cdi_reporte').filter(f => f.endsWith('.jpeg') || f.endsWith('.jpg') || f.endsWith('.png'));
        files.slice(fn => 6).forEach((file, index) => { // limit to 6 for clean formatting or more if needed
            try {
                const imgData = fsObj.readFileSync(`imagenes_cdi_reporte/${file}`);
                paras.push(
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                            new ImageRun({
                                data: imgData,
                                transformation: { width: 350, height: 262 }
                            })
                        ],
                        spacing: { before: 100, after: 50 }
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                            new TextRun({ text: `Figura ${index + 1}: Inspección CEDI D1 Lebrija (${file})`, italics: true, size: 18, color: "666666" })
                        ],
                        spacing: { after: 200 }
                    })
                );
            } catch(e) {}
        });
    }
    return paras;
}

generate();
