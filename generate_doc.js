// Genera "Evaluacion_Riesgos_Incendio_CEDI_D1_Lebrija.docx" a partir del contenido
// del .md y las fotos de imagenes_cdi_reporte/ (formato del "Ejemplo Método Meseri.pdf").
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, ImageRun, Table, TableRow,
  TableCell, WidthType, BorderStyle, ShadingType, VerticalAlign, Header, PageNumber,
  LevelFormat, HeightRule,
} = require('docx');

const FONT = 'Times New Roman';
const FOTOS = path.join(__dirname, 'imagenes_cdi_reporte');
const REFS = path.join(__dirname, 'imagenes_referencia');
const SALIDA = path.join(__dirname, 'Evaluacion_Riesgos_Incendio_CEDI_D1_Lebrija.docx');

// ---------- utilidades ----------
// "27-2" -> "WhatsApp Image 2026-10-22 at 6.20.27 PM (2).jpeg"
function foto(clave) {
  const [seg, n] = clave.split('-');
  return path.join(FOTOS, `WhatsApp Image 2026-10-22 at 6.20.${seg} PM${n ? ` (${n})` : ''}.jpeg`);
}

function jpegSize(buf) {
  let i = 2;
  while (i < buf.length) {
    if (buf[i] !== 0xff) { i++; continue; }
    const m = buf[i + 1];
    const len = buf.readUInt16BE(i + 2);
    if (m >= 0xc0 && m <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(m)) {
      return { h: buf.readUInt16BE(i + 5), w: buf.readUInt16BE(i + 7) };
    }
    i += 2 + len;
  }
  throw new Error('No se pudo leer el tamaño del JPEG');
}

// Imagen ajustada a un ancho/alto máximo (en px a 96 ppp).
function imagen(ruta, maxW, maxH) {
  const data = fs.readFileSync(ruta);
  const { w, h } = jpegSize(data);
  const k = Math.min(maxW / w, maxH / h);
  return new ImageRun({ type: 'jpg', data, transformation: { width: Math.round(w * k), height: Math.round(h * k) } });
}

let nImagen = 0;
// Pie "Imagen N. ..." + foto, para celdas de tabla.
function celdaImagen(ruta, pie, maxW, maxH) {
  nImagen++;
  return [
    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 40 },
      children: [new TextRun({ text: `Imagen ${nImagen}. `, bold: true, size: 18, font: FONT }),
                 new TextRun({ text: pie, size: 18, font: FONT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [imagen(ruta, maxW, maxH)] }),
  ];
}

const p = (text, opts = {}) => new Paragraph({
  alignment: opts.align ?? AlignmentType.JUSTIFIED,
  indent: opts.noIndent ? undefined : { firstLine: 720 },
  spacing: { line: 480, after: 0 },
  children: [new TextRun({ text, font: FONT, size: 24, ...opts.run })],
});

const pRuns = (runs, opts = {}) => new Paragraph({
  alignment: opts.align ?? AlignmentType.JUSTIFIED,
  indent: opts.noIndent ? undefined : { firstLine: 720 },
  spacing: { line: 480, after: 0 },
  children: runs.map(r => (typeof r === 'string' ? new TextRun({ text: r, font: FONT, size: 24 }) : new TextRun({ font: FONT, size: 24, ...r }))),
});

const titulo = (text, opts = {}) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { line: 480, after: 0, before: opts.before ?? 0 },
  pageBreakBefore: opts.pageBreak,
  children: [new TextRun({ text, bold: true, font: FONT, size: 24 })],
});

const vineta = (runs) => new Paragraph({
  numbering: { reference: 'vinetas', level: 0 },
  alignment: AlignmentType.JUSTIFIED,
  spacing: { line: 480, after: 0 },
  children: runs.map(r => (typeof r === 'string' ? new TextRun({ text: r, font: FONT, size: 24 }) : new TextRun({ font: FONT, size: 24, ...r }))),
});

const vacio = (n = 1) => Array.from({ length: n }, () => new Paragraph({ spacing: { line: 480 }, children: [] }));

// ---------- tablas ----------
const borde = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
const bordes = { top: borde, bottom: borde, left: borde, right: borde };
const MARGEN = { top: 70, bottom: 70, left: 90, right: 90 };
const ROSADO = 'FCE4D6';
const AMARILLO = 'FFF2CC';
const GRIS = 'D9D9D9';

function celda(contenido, width, opts = {}) {
  const children = Array.isArray(contenido) && contenido[0] instanceof Paragraph
    ? contenido
    : [new Paragraph({
        alignment: opts.align ?? AlignmentType.LEFT,
        children: [].concat(contenido).map(t => (typeof t === 'string'
          ? new TextRun({ text: t, font: FONT, size: opts.size ?? 20, bold: opts.bold })
          : new TextRun({ font: FONT, size: opts.size ?? 20, ...t }))),
      })];
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: bordes,
    margins: MARGEN,
    verticalAlign: opts.valign ?? VerticalAlign.CENTER,
    columnSpan: opts.span,
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    children,
  });
}

// ---------- Tabla 1: factores propios de las instalaciones (X) ----------
const T1 = [1800, 1400, 3260, 2900]; // = 9360 (6,5")
const IMG1 = { w: 180, h: 215 };

const factoresX = [
  { seccion: '1.1 CONSTRUCCIÓN' },
  { f: '1.1.1 Altura del edificio', c: 1, img: '27', pie: 'Fachada y altura de la nave.',
    j: 'La nave tiene una altura interior libre de 18 m hasta las cerchas que soportan la cubierta (21 m medida por el exterior). Según la tabla del método, para alturas entre 15 y 30 m le corresponde un coeficiente de 1.' },
  { f: '1.1.2 Mayor sector de incendio', c: 0, img: '26', pie: 'Cerramiento lateral de la nave.',
    j: 'La cubierta medida en Google Maps tiene cerca de 17.432 m². No se comprobó sectorización con muros cortafuego, por lo que toda la nave se toma como un solo sector. Al superar los 4.500 m² se asigna un coeficiente de 0.' },
  { f: '1.1.3 Resistencia al fuego', c: 5, img: '29-3', pie: 'Columnas y estructura de cubierta.',
    j: 'La cubierta se apoya en cerchas y correas metálicas sobre columnas de concreto y acero. La estructura metálica se clasifica como no combustible, por lo que se asigna un coeficiente de 5.' },
  { f: '1.1.4 Falsos techos', c: 5, img: '29-1', pie: 'Cubierta interior a la vista.',
    j: 'La cubierta metálica está a la vista en toda la bodega y no tiene cielorraso ni recubrimientos que funcionen como falso techo. Se asigna el coeficiente de 5 (sin falsos techos).' },
  { seccion: '1.2 FACTORES DE SITUACIÓN' },
  { f: '1.2.1 Distancia de los bomberos', c: 8, img: '31', pie: 'Vía de acceso y patio de maniobras.',
    j: 'El Cuerpo de Bomberos Voluntarios de Lebrija atiende emergencias las 24 horas. Se estima una distancia de entre 5 y 10 km por carretera y un tiempo de respuesta de 5 a 10 minutos, por lo que se asigna un coeficiente de 8.' },
  { f: '1.2.2 Accesibilidad del edificio', c: 3, img: '35', pie: 'Portería y acceso al CEDI.',
    j: 'El ingreso es por la vía del sector San Pablo, con portería y talanquera. El patio de maniobras es amplio, pero el predio está en terreno montañoso, elevado sobre su entorno. La accesibilidad se considera media, con coeficiente de 3.' },
  { seccion: '1.3 PROCESOS' },
  { f: '1.3.1 Peligro de activación', c: 5, img: '29-4', pie: 'Estación de carga de montacargas.',
    j: 'No hay llama abierta ni procesos térmicos. Las fuentes de ignición son los montacargas eléctricos en su estación de carga, los tableros eléctricos y la iluminación. El peligro de activación se considera medio, con coeficiente de 5.' },
  { f: '1.3.2 Carga térmica', c: 2, img: '33-3', pie: 'Mercancía almacenada.',
    j: 'Se almacenan alimentos empacados, bebidas, cartón, plástico y estibas de madera en grandes volúmenes. Se estima una carga térmica alta, entre 2.000 y 5.000 MJ/m², a la que corresponde un coeficiente de 2.' },
  { f: '1.3.3 Inflamabilidad', c: 3, img: '28-1', pie: 'Estibas de madera y embalajes.',
    j: 'Predominan sólidos combustibles como cartón, plástico y la madera de las estibas; no se manejan líquidos inflamables en cantidad. La inflamabilidad se clasifica como media, con coeficiente de 3.' },
  { f: '1.3.4 Orden y limpieza', c: 5, img: '29', pie: 'Pasillo de circulación entre racks.',
    j: 'Los pasillos de circulación logística se mantienen despejados, aunque en las zonas de estibas y alistamiento se acumula material. El orden y la limpieza se consideran de nivel medio, con coeficiente de 5.' },
  { f: '1.3.5 Almacenamiento en altura', c: 0, img: '28-2', pie: 'Racks de almacenamiento.',
    j: 'La mercancía se guarda en estanterías tipo rack de varios niveles que superan los 6 m de altura. Según la tabla, para más de 6 m el coeficiente es 0.' },
  { seccion: '1.4 FACTOR DE CONCENTRACIÓN' },
  { f: '1.4.1 Concentración de valores', c: 3, img: '30-1', pie: 'Vista panorámica del almacén.',
    j: 'La mercancía es de consumo masivo y de bajo valor unitario. Repartido sobre la gran superficie del CEDI, el valor se estima por debajo de 600 €/m², por lo que se asigna un coeficiente de 3.' },
  { seccion: '1.5 PROPAGABILIDAD' },
  { f: '1.5.1 En horizontal', c: 3, img: '28-3', pie: 'Pasillos entre bloques de racks.',
    j: 'Los bloques de estantería están separados por pasillos, pero no hay muros cortafuego entre zonas y la mercancía es continua dentro de cada bloque. La propagabilidad horizontal se considera media, con coeficiente de 3.' },
  { f: '1.5.2 En vertical', c: 0, img: '34-2', pie: 'Niveles verticales de los racks.',
    j: 'Los racks de varios niveles con mercancía y estibas de madera forman una columna continua de combustible hasta cerca de la cubierta, lo que facilita la subida del fuego. Se clasifica como alta, con coeficiente de 0.' },
  { seccion: '1.6 DESTRUCTIBILIDAD' },
  { f: '1.6.1 Calor', c: 0, img: '27-2', pie: 'Zona de cava (frío).',
    j: 'La cava de frío y los alimentos empacados se dañan fácilmente con el calor, igual que los paneles y equipos de refrigeración. La destructibilidad por calor es alta, con coeficiente de 0.' },
  { f: '1.6.2 Humo', c: 0, img: '33-2', pie: 'Zona de alimentos secos.',
    j: 'Los productos alimenticios son muy susceptibles a la contaminación por humo; un incendio obligaría a desechar gran parte del inventario. La destructibilidad por humo es alta, con coeficiente de 0.' },
  { f: '1.6.3 Corrosión', c: 5, img: '34-3', pie: 'Bebidas y productos envasados.',
    j: 'Hay productos de aseo, conservas y enlatados que pueden generar o sufrir gases corrosivos, pero no son la mayor parte del inventario. La destructibilidad por corrosión se considera media, con coeficiente de 5.' },
  { f: '1.6.4 Agua', c: 5, img: '28-4', pie: 'Cajas de cartón sobre estibas.',
    j: 'Las cajas de cartón y los empaques se dañan con el agua de los rociadores, aunque muchos productos (bebidas, enlatados, envases plásticos) la resisten. La destructibilidad por agua se considera media, con coeficiente de 5.' },
];
const X = factoresX.filter(r => !r.seccion).reduce((s, r) => s + r.c, 0);

const tabla1 = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: T1,
  rows: [
    new TableRow({ tableHeader: true, children: [celda('FACTORES PROPIOS DE LAS INSTALACIONES', 9360, { span: 4, bold: true, size: 30, align: AlignmentType.CENTER })] }),
    new TableRow({ tableHeader: true, children: ['NOMBRE DEL FACTOR', 'COEFICIENTE', 'JUSTIFICACIÓN', 'IMAGEN'].map((t, i) =>
      celda(t, T1[i], { bold: true, size: 16, align: AlignmentType.CENTER, fill: GRIS })) }),
    ...factoresX.map(r => r.seccion
      ? new TableRow({ cantSplit: true, children: [celda(r.seccion, 9360, { span: 4, bold: true, align: AlignmentType.CENTER, fill: ROSADO })] })
      : new TableRow({ cantSplit: true, children: [
          celda(r.f, T1[0]),
          celda(String(r.c), T1[1], { align: AlignmentType.CENTER, bold: true, size: 22 }),
          celda(r.j, T1[2]),
          celda(celdaImagen(foto(r.img), r.pie, IMG1.w, IMG1.h), T1[3], { valign: VerticalAlign.TOP }),
        ] })),
    new TableRow({ children: [
      celda('TOTAL X', T1[0], { bold: true, fill: AMARILLO }),
      celda(String(X), T1[1], { bold: true, size: 22, align: AlignmentType.CENTER, fill: AMARILLO }),
      celda('Puntuación sobre un máximo de 129 puntos.', T1[2] + T1[3], { span: 2, fill: AMARILLO }),
    ] }),
  ],
});

// ---------- Tabla 2: factores de protección (Y) ----------
const T2 = [1350, 1400, 2410, 2100, 2100]; // = 9360
const IMG2 = { w: 125, h: 160 };

const factoresY = [
  { f: 'Extintores portátiles (EXT)', c: 2, v: 'CV', ref: 'extintor.jpg',
    refPie: 'Extintor multipropósito ABC. Safety Work. Obtenido de https://www.swsafety.com.co/producto/extintor-abc/',
    img: '31-1', pie: 'Extintor portátil en la zona de muelles.',
    j: 'Hay extintores portátiles ubicados en los muelles, la bodega y las áreas operativas, señalizados y con mantenimiento a cargo del personal del CEDI. Con vigilancia se asigna un coeficiente de 2.' },
  { f: 'Bocas de incendio equipadas (BIE)', c: 4, v: 'CV', ref: 'bie.jpg',
    refPie: 'Boca de incendio equipada BIE. Presman. Obtenido de https://www.extintorespresman.es/boca-de-incendio-equipada/',
    img: '30-4', pie: 'Montantes y válvulas de la red contra incendio.',
    j: 'Existe red de tuberías contra incendio con montantes, válvulas de control y gabinetes BIE, alimentada desde el cuarto de bombas. Con vigilancia corresponde un coeficiente de 4.' },
  { f: 'Columnas hidrantes exteriores (CHE)', c: 0, v: 'SV', ref: 'hidrante.jpg',
    refPie: 'Columnas hidrantes exteriores. Extintores Carlisa. Obtenido de https://extintorescarlisa.es/hidrantes-contra-incendios/',
    img: '34', pie: 'Toma exterior junto al muro (por verificar).',
    j: 'Junto al muro exterior se observa una toma de la red contra incendio, pero no se ha confirmado que sea una columna hidrante exterior. Mientras se verifica se asigna 0.' },
  { f: 'Detección automática (DET)', c: 4, v: 'CV', ref: 'detector.jpg',
    refPie: 'Detección automática DET. Asepeyo. Obtenido de https://www.diba.cat/documents/467843/96195101/P1E07011GuiaDeteccion_automatica_incendios_Asepeyo.pdf',
    img: '30-2', pie: 'Dispositivos de detección y alarma bajo cubierta.',
    j: 'Hay detectores de humo y dispositivos de alarma instalados bajo la cubierta y en las áreas operativas. Con vigilancia el coeficiente es 4.' },
  { f: 'Rociadores automáticos (ROC)', c: 8, v: 'CV', ref: 'rociador.jpg',
    refPie: 'Rociadores automáticos ROC. Drexmin. Obtenido de https://drexmin.es/como-funciona-un-rociador-automatico/',
    img: '29-2', pie: 'Red de rociadores bajo la cubierta.',
    j: 'Toda la nave tiene red de tuberías rojas con cabezas de rociadores automáticos bajo la cubierta, válvulas de control y cuarto de bombas. Con vigilancia se asigna un coeficiente de 8.' },
  { f: 'Extinción por agentes gaseosos (IFE)', c: 0, v: 'SV', ref: 'agente_gaseoso.jpg',
    refPie: 'Extinción por agentes gaseosos (IFE). Ruva Seguridad. Obtenido de https://www.extintoresbadalona.com/noticias/tipos-de-agentes-gaseosos/index.html',
    img: null,
    j: 'No hay sistemas de extinción por agentes gaseosos en la zona de abarrotes y bodega; por el tamaño del área este sistema no aplica. Coeficiente de 0.' },
];
const Y = factoresY.reduce((s, r) => s + r.c, 0);

const tabla2 = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: T2,
  rows: [
    new TableRow({ tableHeader: true, children: [celda('FACTORES DE PROTECCIÓN', 9360, { span: 5, bold: true, size: 30, align: AlignmentType.CENTER })] }),
    new TableRow({ tableHeader: true, children: ['NOMBRE DEL FACTOR', 'COEFICIENTE', 'JUSTIFICACIÓN', 'IMAGEN DE REFERENCIA', 'IMAGEN REAL'].map((t, i) =>
      celda(t, T2[i], { bold: true, size: 16, align: AlignmentType.CENTER, fill: GRIS })) }),
    ...factoresY.map(r => new TableRow({ cantSplit: true, children: [
      celda(r.f, T2[0], { align: AlignmentType.CENTER }),
      celda([new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(r.c), bold: true, size: 22, font: FONT })] }),
             new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: r.v === 'CV' ? '(con vigilancia)' : '(sin vigilancia)', size: 16, font: FONT })] })], T2[1]),
      celda(r.j, T2[2]),
      celda(celdaImagen(path.join(REFS, r.ref), r.refPie, IMG2.w, IMG2.h), T2[3], { valign: VerticalAlign.TOP }),
      r.img
        ? celda(celdaImagen(foto(r.img), r.pie, IMG2.w, IMG2.h), T2[4], { valign: VerticalAlign.TOP })
        : celda('NO APLICA', T2[4], { align: AlignmentType.CENTER }),
    ] })),
    new TableRow({ children: [
      celda('TOTAL Y', T2[0], { bold: true, fill: AMARILLO }),
      celda(String(Y), T2[1], { bold: true, size: 22, align: AlignmentType.CENTER, fill: AMARILLO }),
      celda('Puntuación sobre un máximo de 26 puntos.', T2[2] + T2[3] + T2[4], { span: 3, fill: AMARILLO }),
    ] }),
  ],
});

// ---------- cálculo de P ----------
const BCI = 0;
const P = (5 * X) / 129 + (5 * Y) / 26 + BCI;
const coma = (n, d = 2) => n.toFixed(d).replace('.', ',');

const formula = (texto) => new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: texto, font: FONT, size: 24 })] });

const TF = [2000, 5000];
const tablaFormula = new Table({
  width: { size: 7000, type: WidthType.DXA },
  columnWidths: TF,
  alignment: AlignmentType.CENTER,
  rows: [
    ['FÓRMULA', formula('P = (5X / 129) + (5Y / 26) + 1(BCI)')],
    ['PROCEDIMIENTO', formula(`P = 5(${X}) / 129 + 5(${Y}) / 26 + ${BCI}`)],
    ['RESPUESTA', new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `P = ${coma(P)}`, bold: true, font: FONT, size: 26 })] })],
  ].map(([k, v]) => new TableRow({ children: [
    celda(k, TF[0], { bold: true, size: 18, align: AlignmentType.CENTER, fill: GRIS }),
    celda([v], TF[1]),
  ] })),
});

const escenarios = [
  ['Base comprobada (EXT, BIE, DET, ROC)', X, Y, 0],
  ['Con brigada contra incendios certificada', X, Y, 1],
  ['Con hidrantes exteriores y brigada verificados', X, Y + 4, 1],
];
const TE = [3260, 700, 700, 800, 2600, 1300];
const tablaEscenarios = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: TE,
  rows: [
    new TableRow({ tableHeader: true, children: ['Escenario', 'X', 'Y', 'BCI', 'Procedimiento', 'P'].map((t, i) =>
      celda(t, TE[i], { bold: true, align: AlignmentType.CENTER, fill: GRIS })) }),
    ...escenarios.map(([n, x, y, b]) => new TableRow({ children: [
      celda(n, TE[0]),
      celda(String(x), TE[1], { align: AlignmentType.CENTER }),
      celda(String(y), TE[2], { align: AlignmentType.CENTER }),
      celda(String(b), TE[3], { align: AlignmentType.CENTER }),
      celda(`5(${x})/129 + 5(${y})/26 + ${b}`, TE[4], { align: AlignmentType.CENTER }),
      celda(coma((5 * x) / 129 + (5 * y) / 26 + b), TE[5], { align: AlignmentType.CENTER, bold: true }),
    ] })),
  ],
});
const P_brigada = coma((5 * X) / 129 + (5 * Y) / 26 + 1);

// ---------- anexo fotográfico ----------
const anexo = [
  ['31-3', 'Cuarto de bombas: bomba principal contra incendio y tanque de combustible.'],
  ['31-4', 'Cuarto de bombas: vista general del equipo de bombeo.'],
  ['32-1', 'Tubería de succión y descarga de la red contra incendio.'],
  ['32-2', 'Bomba jockey (mantenimiento de presión de la red).'],
  ['32-3', 'Válvulas y tubería principal en el cuarto de bombas.'],
  ['32', 'Tanque hidroneumático y tubería de la red.'],
  ['30-3', 'Tubería de rociadores en el mezanine.'],
  ['28', 'Interior de la cava de frío.'],
  ['27-1', 'Canastas plásticas y estibas en zona de alistamiento.'],
  ['30', 'Zona de servicio con acumulación de estibas.'],
  ['32-4', 'Fachada del bloque de oficinas (mezanine).'],
  ['34-4', 'Muelles de cargue y descargue.'],
  ['33-1', 'Vía perimetral exterior de la nave.'],
  ['34-1', 'Trabajo de mantenimiento en altura bajo la cubierta.'],
];
const TA = [4680, 4680];
const filasAnexo = [];
for (let i = 0; i < anexo.length; i += 2) {
  filasAnexo.push(new TableRow({ cantSplit: true, children: [0, 1].map(k => {
    const a = anexo[i + k];
    return a ? celda(celdaImagen(foto(a[0]), a[1], 280, 250), TA[k], { valign: VerticalAlign.TOP }) : celda('', TA[k]);
  }) }));
}
const tablaAnexo = new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: TA, rows: filasAnexo });

// ---------- carátula ----------
const caratula = [
  ...vacio(1),
  new Table({
    width: { size: 3600, type: WidthType.DXA },
    columnWidths: [3600],
    alignment: AlignmentType.CENTER,
    rows: [new TableRow({ height: { value: 1800, rule: HeightRule.EXACT }, children: [new TableCell({
      width: { size: 3600, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      borders: { top: { style: BorderStyle.DASHED, size: 8, color: '999999' }, bottom: { style: BorderStyle.DASHED, size: 8, color: '999999' },
                 left: { style: BorderStyle.DASHED, size: 8, color: '999999' }, right: { style: BorderStyle.DASHED, size: 8, color: '999999' } },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Espacio para el logo', italics: true, color: '888888', font: FONT, size: 22 })] })],
    })] })],
  }),
  ...vacio(2),
  titulo('Evaluación de Riesgos de Incendio'),
  titulo('Método de Meseri'),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 480 }, children: [new TextRun({ text: 'CEDI D1 Bucaramanga / Lebrija', font: FONT, size: 24 })] }),
  ...vacio(3),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 480 }, children: [new TextRun({ text: '[Nombre del alumno]', bold: true, font: FONT, size: 24, highlight: 'yellow' })] }),
  ...vacio(1),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 480 }, children: [
    new TextRun({ text: 'Profesor: ', bold: true, font: FONT, size: 24 }),
    new TextRun({ text: '[Nombre del profesor]', font: FONT, size: 24, highlight: 'yellow' })] }),
  ...vacio(4),
  ...['Unidades Tecnológicas de Santander', 'Electiva Complementaria Salud Ocupacional', 'Octubre de 2026', 'Bucaramanga, Santander']
    .map(t => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 480 }, children: [new TextRun({ text: t, font: FONT, size: 24 })] })),
];

// ---------- cuerpo ----------
const introduccion = [
  titulo('Introducción', { pageBreak: true }),
  p('El Centro de Distribución (CEDI) evaluado está ubicado en el área rural de Lebrija (Santander), sector San Pablo, sobre la vía Bucaramanga–Barrancabermeja, a unos 15 minutos de Bucaramanga y a cerca de 2,5 km del aeropuerto Palonegro. Hace parte de la red logística de Tiendas D1 (Koba Colombia S.A.S.), cadena de tiendas de descuento que contrata personal de mantenimiento para sus centros de distribución en Santander.'),
  p('Se trata de una bodega de distribución de gran tamaño, con una cubierta de aproximadamente 17.432 m² según la medición hecha en Google Maps. La nave tiene una altura interior libre de 18 m (21 m por el exterior), muros de bloque de concreto en la parte baja, cerramiento y cubierta en lámina metálica, cerchas metálicas y columnas de concreto y acero. Cuenta con una zona principal de almacenamiento en estanterías tipo rack de varios niveles, áreas de cava (frío), muelles de cargue y descargue, zonas de servicio para la operación y un mezanine con oficinas, cafetería y servicios para los empleados.'),
  p('La actividad del CEDI consiste en recibir, almacenar y despachar alimentos empacados, bebidas, productos de aseo, conservas y enlatados hacia las tiendas de la región. La mercancía se mueve sobre estibas de madera y canastas plásticas con montacargas eléctricos, que se recargan en una estación de carga dentro de la bodega. No hay procesos con llama abierta ni manejo de líquidos inflamables en cantidad, pero el volumen de cartón, plástico y madera almacenado es muy alto.'),
  p('En cuanto a la protección contra incendios, el CEDI cuenta con un cuarto de bombas con bomba principal, bomba jockey y tanque hidroneumático, que alimenta una red de tuberías de color rojo con montantes, válvulas de control, gabinetes de bocas de incendio equipadas (BIE) y rociadores automáticos instalados bajo toda la cubierta. Además, hay extintores portátiles en los muelles y áreas operativas y dispositivos de detección y alarma bajo la cubierta.'),
  p('Para evaluar el riesgo de incendio se aplicó el Método de Meseri, propuesto por la Fundación MAPFRE. El método valora por separado los factores propios de las instalaciones, que generan o agravan el riesgo (X, máximo 129 puntos), y los factores de protección, que lo reducen (Y, máximo 26 puntos). Con estos valores se calcula el coeficiente de protección P, que va de 0 a 10; un resultado igual o superior a 5 se considera aceptable. Los coeficientes se asignaron a partir de la inspección fotográfica de las instalaciones realizada entre marzo y octubre de 2026.'),
];

const evaluacion = [
  titulo('Evaluación de Riesgos de Incendio – Método de Meseri', { before: 240 }),
  new Paragraph({ spacing: { before: 240 }, children: [new TextRun({ text: 'Tabla 1.', bold: true, font: FONT, size: 24 })] }),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: 'Factores Propios de las Instalaciones', italics: true, font: FONT, size: 24 })] }),
  tabla1,
  new Paragraph({ pageBreakBefore: true, children: [new TextRun({ text: 'Tabla 2.', bold: true, font: FONT, size: 24 })] }),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: 'Factores de Protección', italics: true, font: FONT, size: 24 })] }),
  tabla2,
  new Paragraph({ spacing: { before: 120 }, children: [new TextRun({ text: 'CV: con vigilancia. SV: sin vigilancia.', italics: true, font: FONT, size: 18 })] }),
];

const calculo = [
  titulo('Coeficiente de Protección contra Incendios (P)', { pageBreak: true }),
  p('Con los totales de las dos tablas se aplica la fórmula del Método de Meseri, donde X es la suma de los factores propios de las instalaciones, Y la suma de los factores de protección y BCI vale 1 si existe una brigada contra incendios y 0 si no existe. Como la brigada del CEDI no se ha podido confirmar, se toma BCI = 0.'),
  ...vacio(1),
  tablaFormula,
  ...vacio(1),
  pRuns([`P = (5 × ${X} / 129) + (5 × ${Y} / 26) + ${BCI} = ${coma((5 * X) / 129, 3)} + ${coma((5 * Y) / 26, 3)} + ${BCI} = `, { text: coma(P), bold: true }], { align: AlignmentType.CENTER, noIndent: true }),
  ...vacio(1),
  p('La siguiente tabla compara el resultado actual con dos escenarios de mejora que dependen de factores aún por verificar:'),
  new Paragraph({ spacing: { before: 240 }, children: [new TextRun({ text: 'Tabla 3.', bold: true, font: FONT, size: 24 })] }),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: 'Escenarios del Coeficiente de Protección', italics: true, font: FONT, size: 24 })] }),
  tablaEscenarios,
];

const conclusion = [
  titulo('Conclusión', { pageBreak: true }),
  p(`Este trabajo aplicó el Método de Meseri para evaluar el riesgo de incendio del CEDI D1 de Lebrija. Con base en la inspección fotográfica se calificaron los factores propios de las instalaciones, que sumaron X = ${X} de 129 puntos, y los factores de protección, que sumaron Y = ${Y} de 26 puntos. Al aplicar la fórmula se obtuvo un coeficiente de protección P = ${coma(P)}.`),
  p('Según la escala del Método de Meseri de la Fundación MAPFRE, el riesgo se califica así:'),
  vineta([{ text: 'Menor de 3: ', bold: true }, 'muy malo.']),
  vineta([{ text: 'De 3 a 5: ', bold: true }, 'malo.']),
  vineta([{ text: 'De 5 a 8: ', bold: true }, 'bueno.']),
  vineta([{ text: 'Mayor de 8: ', bold: true }, 'muy bueno.']),
  pRuns([`Por lo tanto, la puntuación de `, { text: coma(P), bold: true }, ' clasifica el riesgo global del CEDI como ', { text: 'bueno', bold: true }, ' y supera el valor mínimo aceptable de 5.']),
  p('Las instalaciones por sí solas son desfavorables: el CEDI es un único sector de más de 4.500 m² sin muros cortafuego, almacena mercancía en racks de más de 6 m que facilitan la propagación vertical, tiene una carga térmica alta y sus productos alimenticios se dañan con facilidad por el calor y el humo. Lo que eleva el resultado son los sistemas de protección, sobre todo los rociadores automáticos (8 puntos), la detección automática (4 puntos) y la red de BIE (4 puntos), que compensan en buena medida esa vulnerabilidad.'),
  p('Para mantener y mejorar este nivel de protección se recomienda:'),
  vineta(['Mantener el programa de inspección, prueba y mantenimiento periódico de la red de rociadores, el cuarto de bombas y las válvulas de control.']),
  vineta([`Conformar o certificar formalmente la brigada contra incendios, lo que suma 1 punto (BCI = 1) y eleva el coeficiente a ${P_brigada}.`]),
  vineta(['Verificar si la toma exterior observada junto al muro corresponde a un hidrante; de confirmarse, el factor CHE sumaría puntos adicionales.']),
  vineta(['Evaluar la sectorización física con muros cortafuego en futuras ampliaciones y controlar la acumulación de estibas de madera en las zonas de servicio.']),
  vineta(['Capacitar al personal en el uso de extintores y en los procedimientos de evacuación.']),
  p('En resumen, la aplicación del Método de Meseri permitió identificar tanto las debilidades propias de un centro de distribución de gran tamaño como el aporte de sus sistemas de protección. El ejercicio muestra la importancia de mantener estos sistemas en buen estado, ya que de ellos depende que el riesgo se mantenga en un nivel aceptable para los trabajadores, la mercancía y la continuidad de la operación.'),
];

const referencia = (runs) => new Paragraph({
  indent: { left: 720, hanging: 720 },
  spacing: { line: 480, after: 0 },
  children: runs.map(r => (typeof r === 'string' ? new TextRun({ text: r, font: FONT, size: 24 }) : new TextRun({ font: FONT, size: 24, ...r }))),
});

const referencias = [
  titulo('Referencias', { pageBreak: true }),
  referencia(['Cuerpo de Bomberos Voluntarios de Lebrija. (s. f.). ', { text: 'Atención de emergencias', italics: true }, '. Lebrija, Santander.']),
  referencia(['Drexmin. (s. f.). ', { text: 'Cómo funciona un rociador automático', italics: true }, '. https://drexmin.es/como-funciona-un-rociador-automatico/']),
  referencia(['Extintores Carlisa. (s. f.). ', { text: 'Hidrantes contra incendios', italics: true }, '. https://extintorescarlisa.es/hidrantes-contra-incendios/']),
  referencia(['Extintores Presman. (s. f.). ', { text: 'Boca de incendio equipada', italics: true }, '. https://www.extintorespresman.es/boca-de-incendio-equipada/']),
  referencia(['Fondo Inmobiliario Colombia. (s. f.). ', { text: 'Ficha CEDI Bucaramanga / Lebrija', italics: true }, '.']),
  referencia(['Fundación MAPFRE. (s. f.). ', { text: 'Método simplificado de evaluación del riesgo de incendio: MESERI', italics: true }, '. Fundación MAPFRE.']),
  referencia(['Koba Colombia S.A.S. (s. f.). ', { text: 'Tiendas D1', italics: true }, '.']),
  referencia(['Asepeyo. (s. f.). ', { text: 'Guía de detección automática de incendios', italics: true }, '. https://www.diba.cat/documents/467843/96195101/P1E07011GuiaDeteccion_automatica_incendios_Asepeyo.pdf']),
  referencia(['Ruva Seguridad. (s. f.). ', { text: 'Tipos de agentes gaseosos', italics: true }, '. https://www.extintoresbadalona.com/noticias/tipos-de-agentes-gaseosos/index.html']),
  referencia(['Safety Work. (s. f.). ', { text: 'Extintor ABC', italics: true }, '. https://www.swsafety.com.co/producto/extintor-abc/']),
];

const anexoSec = [
  titulo('Anexo. Registro Fotográfico Complementario', { pageBreak: true }),
  new Paragraph({ spacing: { after: 120 }, children: [] }),
  tablaAnexo,
];

// ---------- documento ----------
const doc = new Document({
  creator: 'CEDI D1 Lebrija',
  title: 'Evaluación de Riesgos de Incendio – Método de Meseri',
  styles: { default: { document: { run: { font: FONT, size: 24 } } } },
  numbering: { config: [{ reference: 'vinetas', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 24 })] })] }) },
    children: [...caratula, ...introduccion, ...evaluacion, ...calculo, ...conclusion, ...referencias, ...anexoSec],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(SALIDA, buf);
  console.log(`OK: ${SALIDA} | X=${X} Y=${Y} P=${coma(P)} | imágenes=${nImagen}`);
});
