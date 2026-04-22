const ExcelJS = require('exceljs');

async function createSimulator() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Simulador PMM';
  workbook.created = new Date();

  // ===== COLOR PALETTE =====
  const C = {
    teal: '00838F', tealDark: '004D40', tealLight: 'B2DFDB', tealVLight: 'E0F7FA',
    white: 'FFFFFF', black: '000000', red: 'F44336', redLight: 'FFCDD2',
    green: '4CAF50', greenLight: 'C8E6C9', amber: 'FFC107', amberLight: 'FFF8E1',
    gray: 'ECEFF1', grayDark: '607D8B'
  };

  // ===== FONT HELPERS =====
  const fontTitle = { name: 'Calibri', size: 14, bold: true, color: { argb: C.white } };
  const fontSection = { name: 'Calibri', size: 11, bold: true, color: { argb: C.white } };
  const fontHeader = { name: 'Calibri', size: 10, bold: true, color: { argb: C.white } };
  const fontNormal = { name: 'Calibri', size: 10, color: { argb: C.black } };
  const fontBold = { name: 'Calibri', size: 10, bold: true, color: { argb: C.black } };
  const fontBoldWhite = { name: 'Calibri', size: 10, bold: true, color: { argb: C.white } };
  const fontLabel = { name: 'Calibri', size: 11, bold: true, color: { argb: C.tealDark } };

  const fillSolid = (color) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: color } });
  const borderThin = { top: { style: 'thin', color: { argb: C.tealLight } }, bottom: { style: 'thin', color: { argb: C.tealLight } }, left: { style: 'thin', color: { argb: C.tealLight } }, right: { style: 'thin', color: { argb: C.tealLight } } };
  const alignCenter = { horizontal: 'center', vertical: 'middle' };
  const alignRight = { horizontal: 'right', vertical: 'middle' };
  const alignLeft = { horizontal: 'left', vertical: 'middle' };
  const numFmtEur = '#,##0€';
  const numFmtPct = '0.0%';

  // ===== PRODUCTS & COMPANIES =====
  const products = [
    { name: 'Créditos', sublimit: 2000000 },
    { name: 'Descuento Comercial', sublimit: 8000000 },
    { name: 'Confirming', sublimit: 8000000 },
    { name: 'Comex', sublimit: 8000000 },
    { name: 'Préstamos', sublimit: 4000000 },
  ];
  const companies = ['Empresa 1', 'Empresa 2', 'Empresa 3', 'Empresa 4'];
  const globalLimit = 8000000;

  // Sample data (respects all limits)
  const sampleData = [
    [500000, 300000, 200000, 400000],   // Créditos = 1.4M < 2M
    [1000000, 500000, 800000, 700000],  // Dto.Com = 3M < 8M
    [300000, 200000, 400000, 100000],   // Confirming = 1M < 8M
    [0, 100000, 200000, 0],            // Comex = 0.3M < 8M
    [500000, 300000, 400000, 300000],   // Préstamos = 1.5M < 4M
  ]; // Total global = 7.2M < 8M

  // ===== WORKSHEET =====
  const ws = workbook.addWorksheet('Simulador PMM', {
    properties: { tabColor: { argb: C.teal } },
    views: [{ showGridLines: false }]
  });

  // Column widths: A=Producto, B=Sublímite, C-F=Empresas, G=TotalDisp, H=Disponible, I=%Util
  const colWidths = [24, 20, 18, 18, 18, 18, 20, 20, 14];
  colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // Helper to style a cell
  function styleCell(cell, opts = {}) {
    if (opts.font) cell.font = opts.font;
    if (opts.fill) cell.fill = opts.fill;
    if (opts.align) cell.alignment = opts.align;
    if (opts.border) cell.border = opts.border;
    if (opts.numFmt) cell.numFmt = opts.numFmt;
    if (opts.protection) cell.protection = opts.protection;
  }

  // Helper: style a range of cells in a row
  function styleRow(row, startCol, endCol, opts) {
    for (let c = startCol; c <= endCol; c++) {
      styleCell(ws.getRow(row).getCell(c), opts);
    }
  }

  // ============================================================
  // ROW 1: TITLE
  // ============================================================
  ws.mergeCells('A1:I1');
  styleCell(ws.getCell('A1'), { font: fontTitle, fill: fillSolid(C.teal), align: alignCenter });
  ws.getCell('A1').value = 'SIMULADOR PÓLIZA MULTIPRODUCTO MULTIEMPRESA (PMM)';
  ws.getRow(1).height = 38;

  // ============================================================
  // ROW 3: GLOBAL LIMIT CONFIG
  // ============================================================
  ws.mergeCells('A3:B3');
  styleCell(ws.getCell('A3'), { font: fontLabel, align: alignRight });
  ws.getCell('A3').value = 'LÍMITE GLOBAL PMM:';
  styleCell(ws.getCell('C3'), {
    font: { name: 'Calibri', size: 14, bold: true, color: { argb: C.teal } },
    fill: fillSolid(C.tealVLight), align: alignCenter, border: borderThin, numFmt: numFmtEur,
    protection: { locked: false }
  });
  ws.getCell('C3').value = globalLimit;
  ws.getRow(3).height = 30;

  // ============================================================
  // ROW 5: SECTION HEADER - MATRIZ
  // ============================================================
  ws.mergeCells('A5:I5');
  styleCell(ws.getCell('A5'), { font: fontSection, fill: fillSolid(C.tealDark), align: alignCenter });
  ws.getCell('A5').value = 'MATRIZ DE DISPOSICIONES POR EMPRESA Y PRODUCTO';
  ws.getRow(5).height = 28;

  // ============================================================
  // ROW 6: TABLE HEADERS
  // ============================================================
  const headers = ['Tipo de Producto', 'Sublímite', ...companies, 'Total Dispuesto', 'Disponible', '% Utilización'];
  headers.forEach((h, i) => {
    const cell = ws.getRow(6).getCell(i + 1);
    cell.value = h;
    styleCell(cell, { font: fontHeader, fill: fillSolid(C.teal), align: alignCenter, border: borderThin });
  });
  ws.getRow(6).height = 28;

  // ============================================================
  // ROWS 7-11: PRODUCT DATA ROWS
  // ============================================================
  const firstProdRow = 7;
  const lastProdRow = firstProdRow + products.length - 1; // 11
  const firstCompCol = 3; // C
  const lastCompCol = firstCompCol + companies.length - 1; // F
  const totalDispCol = lastCompCol + 1; // G
  const disponibleCol = totalDispCol + 1; // H
  const pctCol = disponibleCol + 1; // I

  // Column letters helper
  const colLetter = (n) => String.fromCharCode(64 + n);

  products.forEach((prod, pi) => {
    const r = firstProdRow + pi;
    const row = ws.getRow(r);
    row.height = 24;

    // A: Product name
    styleCell(row.getCell(1), { font: fontBold, fill: fillSolid(pi % 2 === 0 ? C.gray : C.white), align: alignLeft, border: borderThin });
    row.getCell(1).value = prod.name;

    // B: Sublimit (editable)
    styleCell(row.getCell(2), {
      font: fontBold, fill: fillSolid(C.tealVLight), align: alignCenter, border: borderThin,
      numFmt: numFmtEur, protection: { locked: false }
    });
    row.getCell(2).value = prod.sublimit;

    // C-F: Company dispositions (editable input cells)
    companies.forEach((_, ci) => {
      const c = firstCompCol + ci;
      styleCell(row.getCell(c), {
        font: fontNormal, fill: fillSolid(C.amberLight), align: alignCenter, border: borderThin,
        numFmt: numFmtEur, protection: { locked: false }
      });
      row.getCell(c).value = sampleData[pi][ci];
    });

    // G: Total Dispuesto = SUM(C:F)
    const cStart = colLetter(firstCompCol);
    const cEnd = colLetter(lastCompCol);
    styleCell(row.getCell(totalDispCol), {
      font: fontBold, fill: fillSolid(C.tealLight), align: alignCenter, border: borderThin, numFmt: numFmtEur
    });
    row.getCell(totalDispCol).value = { formula: `SUM(${cStart}${r}:${cEnd}${r})` };

    // H: Disponible = MIN(Sublímite - Dispuesto, MAX(0, LímiteGlobal - TotalDispGlobal))
    const totalRow = lastProdRow + 1; // 12
    const gCol = colLetter(totalDispCol);
    styleCell(row.getCell(disponibleCol), {
      font: fontBold, fill: fillSolid(C.greenLight), align: alignCenter, border: borderThin, numFmt: numFmtEur
    });
    row.getCell(disponibleCol).value = {
      formula: `MIN(B${r}-${gCol}${r}, MAX(0, $C$3-${gCol}$${totalRow}))`
    };

    // I: % Utilización = Total Dispuesto / Sublímite
    styleCell(row.getCell(pctCol), {
      font: fontBold, fill: fillSolid(pi % 2 === 0 ? C.gray : C.white), align: alignCenter, border: borderThin, numFmt: numFmtPct
    });
    row.getCell(pctCol).value = {
      formula: `IF(B${r}=0,0,${gCol}${r}/B${r})`
    };
  });

  // ============================================================
  // ROW 12: TOTALS ROW
  // ============================================================
  const totalRow = lastProdRow + 1; // 12
  const tRow = ws.getRow(totalRow);
  tRow.height = 28;

  // A: Label
  styleCell(tRow.getCell(1), { font: fontBoldWhite, fill: fillSolid(C.tealDark), align: alignCenter, border: borderThin });
  tRow.getCell(1).value = 'TOTAL PMM';

  // B: Sum of sublimits (informational)
  styleCell(tRow.getCell(2), { font: fontBoldWhite, fill: fillSolid(C.tealDark), align: alignCenter, border: borderThin, numFmt: numFmtEur });
  tRow.getCell(2).value = { formula: `SUM(B${firstProdRow}:B${lastProdRow})` };

  // C-F: Sum per company
  companies.forEach((_, ci) => {
    const c = firstCompCol + ci;
    const cl = colLetter(c);
    styleCell(tRow.getCell(c), { font: fontBoldWhite, fill: fillSolid(C.tealDark), align: alignCenter, border: borderThin, numFmt: numFmtEur });
    tRow.getCell(c).value = { formula: `SUM(${cl}${firstProdRow}:${cl}${lastProdRow})` };
  });

  // G: Total Dispuesto Global
  const gCol = colLetter(totalDispCol);
  styleCell(tRow.getCell(totalDispCol), { font: fontBoldWhite, fill: fillSolid(C.tealDark), align: alignCenter, border: borderThin, numFmt: numFmtEur });
  tRow.getCell(totalDispCol).value = { formula: `SUM(${gCol}${firstProdRow}:${gCol}${lastProdRow})` };

  // H: Disponible Global = Límite Global - Total Dispuesto
  const hCol = colLetter(disponibleCol);
  styleCell(tRow.getCell(disponibleCol), { font: fontBoldWhite, fill: fillSolid(C.tealDark), align: alignCenter, border: borderThin, numFmt: numFmtEur });
  tRow.getCell(disponibleCol).value = { formula: `$C$3-${gCol}${totalRow}` };

  // I: % Utilización Global
  styleCell(tRow.getCell(pctCol), { font: fontBoldWhite, fill: fillSolid(C.tealDark), align: alignCenter, border: borderThin, numFmt: numFmtPct });
  tRow.getCell(pctCol).value = { formula: `IF($C$3=0,0,${gCol}${totalRow}/$C$3)` };

  // ============================================================
  // ROW 14: SECTION HEADER - INDICADORES GLOBALES
  // ============================================================
  ws.mergeCells('A14:E14');
  styleCell(ws.getCell('A14'), { font: fontSection, fill: fillSolid(C.tealDark), align: alignCenter });
  ws.getCell('A14').value = 'INDICADORES GLOBALES';
  ws.getRow(14).height = 28;

  // ROW 15: Headers
  const indHeaders = ['Indicador', 'Importe', '% sobre Límite', 'Estado', ''];
  indHeaders.forEach((h, i) => {
    const cell = ws.getRow(15).getCell(i + 1);
    cell.value = h;
    styleCell(cell, { font: fontHeader, fill: fillSolid(C.teal), align: alignCenter, border: borderThin });
  });

  // ROW 16: Límite Global
  ws.getRow(16).height = 24;
  styleCell(ws.getCell('A16'), { font: fontBold, align: alignLeft, border: borderThin, fill: fillSolid(C.gray) });
  ws.getCell('A16').value = 'Límite Global PMM';
  styleCell(ws.getCell('B16'), { font: fontBold, align: alignCenter, border: borderThin, numFmt: numFmtEur, fill: fillSolid(C.gray) });
  ws.getCell('B16').value = { formula: '$C$3' };
  styleCell(ws.getCell('C16'), { font: fontBold, align: alignCenter, border: borderThin, numFmt: numFmtPct, fill: fillSolid(C.gray) });
  ws.getCell('C16').value = 1;
  styleCell(ws.getCell('D16'), { font: fontBold, align: alignCenter, border: borderThin, fill: fillSolid(C.gray) });
  ws.getCell('D16').value = '—';

  // ROW 17: Total Dispuesto
  ws.getRow(17).height = 24;
  styleCell(ws.getCell('A17'), { font: fontBold, align: alignLeft, border: borderThin });
  ws.getCell('A17').value = 'Total Dispuesto';
  styleCell(ws.getCell('B17'), { font: fontBold, align: alignCenter, border: borderThin, numFmt: numFmtEur });
  ws.getCell('B17').value = { formula: `${gCol}${totalRow}` };
  styleCell(ws.getCell('C17'), { font: fontBold, align: alignCenter, border: borderThin, numFmt: numFmtPct });
  ws.getCell('C17').value = { formula: `IF($C$3=0,0,B17/$C$3)` };
  styleCell(ws.getCell('D17'), { font: fontBold, align: alignCenter, border: borderThin });
  ws.getCell('D17').value = { formula: `IF(B17>$C$3,"⚠ EXCEDIDO","✓ OK")` };

  // ROW 18: Disponible Global
  ws.getRow(18).height = 24;
  styleCell(ws.getCell('A18'), { font: fontBold, align: alignLeft, border: borderThin, fill: fillSolid(C.greenLight) });
  ws.getCell('A18').value = 'Disponible Global';
  styleCell(ws.getCell('B18'), { font: fontBold, align: alignCenter, border: borderThin, numFmt: numFmtEur, fill: fillSolid(C.greenLight) });
  ws.getCell('B18').value = { formula: `$C$3-B17` };
  styleCell(ws.getCell('C18'), { font: fontBold, align: alignCenter, border: borderThin, numFmt: numFmtPct, fill: fillSolid(C.greenLight) });
  ws.getCell('C18').value = { formula: `IF($C$3=0,0,B18/$C$3)` };
  styleCell(ws.getCell('D18'), { font: fontBold, align: alignCenter, border: borderThin, fill: fillSolid(C.greenLight) });
  ws.getCell('D18').value = { formula: `IF(B18<0,"⚠ SIN DISPONIBLE","✓ DISPONIBLE")` };

  // ============================================================
  // ROW 20: SECTION - INDICADORES POR PRODUCTO
  // ============================================================
  ws.mergeCells('A20:I20');
  styleCell(ws.getCell('A20'), { font: fontSection, fill: fillSolid(C.tealDark), align: alignCenter });
  ws.getCell('A20').value = 'INDICADORES POR PRODUCTO';
  ws.getRow(20).height = 28;

  // ROW 21: Headers
  const prodIndHeaders = ['Producto', 'Sublímite', 'Dispuesto', 'Disponible', '% Sublímite', '% Global', 'Estado Sublímite', 'Estado Global', 'Margen Global'];
  prodIndHeaders.forEach((h, i) => {
    const cell = ws.getRow(21).getCell(i + 1);
    cell.value = h;
    styleCell(cell, { font: fontHeader, fill: fillSolid(C.teal), align: alignCenter, border: borderThin });
  });

  // ROWS 22-26: Per-product indicators
  products.forEach((prod, pi) => {
    const r = 22 + pi;
    const prodDataRow = firstProdRow + pi;
    const row = ws.getRow(r);
    row.height = 24;
    const bg = pi % 2 === 0 ? C.gray : C.white;

    // A: Producto
    styleCell(row.getCell(1), { font: fontBold, align: alignLeft, border: borderThin, fill: fillSolid(bg) });
    row.getCell(1).value = prod.name;
    // B: Sublímite (reference)
    styleCell(row.getCell(2), { font: fontNormal, align: alignCenter, border: borderThin, numFmt: numFmtEur, fill: fillSolid(bg) });
    row.getCell(2).value = { formula: `B${prodDataRow}` };
    // C: Dispuesto
    styleCell(row.getCell(3), { font: fontNormal, align: alignCenter, border: borderThin, numFmt: numFmtEur, fill: fillSolid(bg) });
    row.getCell(3).value = { formula: `${gCol}${prodDataRow}` };
    // D: Disponible
    styleCell(row.getCell(4), { font: fontBold, align: alignCenter, border: borderThin, numFmt: numFmtEur, fill: fillSolid(C.greenLight) });
    row.getCell(4).value = { formula: `${hCol}${prodDataRow}` };
    // E: % Sublímite
    styleCell(row.getCell(5), { font: fontNormal, align: alignCenter, border: borderThin, numFmt: numFmtPct, fill: fillSolid(bg) });
    row.getCell(5).value = { formula: `IF(B${r}=0,0,C${r}/B${r})` };
    // F: % Global
    styleCell(row.getCell(6), { font: fontNormal, align: alignCenter, border: borderThin, numFmt: numFmtPct, fill: fillSolid(bg) });
    row.getCell(6).value = { formula: `IF($C$3=0,0,C${r}/$C$3)` };
    // G: Estado Sublímite
    styleCell(row.getCell(7), { font: fontBold, align: alignCenter, border: borderThin, fill: fillSolid(bg) });
    row.getCell(7).value = { formula: `IF(C${r}>B${r},"⚠ EXCEDIDO",IF(C${r}/B${r}>0.8,"⚡ ALTO","✓ OK"))` };
    // H: Estado Global
    styleCell(row.getCell(8), { font: fontBold, align: alignCenter, border: borderThin, fill: fillSolid(bg) });
    row.getCell(8).value = { formula: `IF(${gCol}${totalRow}>$C$3,"⚠ EXCEDIDO","✓ OK")` };
    // I: Margen Global restante
    styleCell(row.getCell(9), { font: fontNormal, align: alignCenter, border: borderThin, numFmt: numFmtEur, fill: fillSolid(bg) });
    row.getCell(9).value = { formula: `MAX(0,$C$3-${gCol}${totalRow})` };
  });

  // ============================================================
  // CONDITIONAL FORMATTING
  // ============================================================
  // Red fill when disponible < 0 in product rows
  ws.addConditionalFormatting({
    ref: `${hCol}${firstProdRow}:${hCol}${lastProdRow}`,
    rules: [{ type: 'cellIs', operator: 'lessThan', formulae: [0], style: { fill: fillSolid(C.redLight), font: { color: { argb: C.red }, bold: true } }, priority: 1 }]
  });

  // Red fill when global disponible < 0
  ws.addConditionalFormatting({
    ref: `${hCol}${totalRow}`,
    rules: [{ type: 'cellIs', operator: 'lessThan', formulae: [0], style: { fill: fillSolid(C.redLight), font: { color: { argb: C.red }, bold: true } }, priority: 2 }]
  });

  // Amber when % utilización > 80%
  ws.addConditionalFormatting({
    ref: `${colLetter(pctCol)}${firstProdRow}:${colLetter(pctCol)}${lastProdRow}`,
    rules: [
      { type: 'cellIs', operator: 'greaterThan', formulae: [1], style: { fill: fillSolid(C.redLight), font: { color: { argb: C.red }, bold: true } }, priority: 3 },
      { type: 'cellIs', operator: 'greaterThan', formulae: [0.8], style: { fill: fillSolid('FFF3E0'), font: { color: { argb: 'E65100' }, bold: true } }, priority: 4 }
    ]
  });

  // ============================================================
  // DATA VALIDATION - prevent negative numbers in input cells
  // ============================================================
  for (let r = firstProdRow; r <= lastProdRow; r++) {
    for (let c = firstCompCol; c <= lastCompCol; c++) {
      ws.getCell(r, c).dataValidation = {
        type: 'decimal', operator: 'greaterThanOrEqual',
        formulae: [0], showErrorMessage: true,
        errorTitle: 'Valor no válido', error: 'Las disposiciones no pueden ser negativas.'
      };
    }
  }

  // ============================================================
  // PRINT SETUP
  // ============================================================
  ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

  // ============================================================
  // INSTRUCTIONS SHEET
  // ============================================================
  const wsHelp = workbook.addWorksheet('Instrucciones', { properties: { tabColor: { argb: C.grayDark } } });
  wsHelp.getColumn(1).width = 60;
  wsHelp.getColumn(2).width = 40;

  const instructions = [
    ['INSTRUCCIONES DE USO - SIMULADOR PMM', ''],
    ['', ''],
    ['1. LÍMITE GLOBAL', 'Celda C3 en hoja "Simulador PMM"'],
    ['   Modifique el valor del límite global de la póliza paraguas.', ''],
    ['', ''],
    ['2. SUBLÍMITES POR PRODUCTO', 'Columna B, filas 7-11'],
    ['   Modifique los sublímites de cada tipo de producto.', ''],
    ['   Los sublímites pueden superar el límite global (son techos por producto).', ''],
    ['', ''],
    ['3. DISPOSICIONES POR EMPRESA', 'Columnas C-F, filas 7-11 (celdas amarillas)'],
    ['   Introduzca los importes dispuestos por cada empresa en cada producto.', ''],
    ['   Las celdas amarillas son editables. No se permiten valores negativos.', ''],
    ['', ''],
    ['4. INDICADORES AUTOMÁTICOS', ''],
    ['   - Total Dispuesto: suma de disposiciones por producto y empresa.', ''],
    ['   - Disponible: MIN(sublímite - dispuesto, límite global - total global).', ''],
    ['   - % Utilización: porcentaje del sublímite consumido.', ''],
    ['   - El disponible nunca permite superar el límite global.', ''],
    ['', ''],
    ['5. ALERTAS', ''],
    ['   - Rojo: sublímite o límite global excedido / disponible negativo.', ''],
    ['   - Ámbar: utilización > 80%.', ''],
    ['   - ⚠ EXCEDIDO: se ha superado un límite.', ''],
    ['   - ✓ OK: dentro de parámetros.', ''],
    ['', ''],
    ['6. REGLA CLAVE', ''],
    ['   El límite global NO puede superarse en ningún momento.', ''],
    ['   El disponible por producto se ajusta automáticamente para', ''],
    ['   respetar tanto el sublímite del producto como el límite global.', ''],
  ];

  instructions.forEach((row, i) => {
    wsHelp.getRow(i + 1).getCell(1).value = row[0];
    wsHelp.getRow(i + 1).getCell(2).value = row[1];
    if (i === 0) {
      wsHelp.getRow(1).getCell(1).font = fontTitle;
      wsHelp.getRow(1).getCell(1).font = { ...fontTitle, color: { argb: C.tealDark } };
    }
  });

  // ============================================================
  // SAVE
  // ============================================================
  const filePath = 'Simulador_PMM.xlsx';
  await workbook.xlsx.writeFile(filePath);
  console.log(`✅ Archivo generado: ${filePath}`);
}

createSimulator().catch(console.error);
