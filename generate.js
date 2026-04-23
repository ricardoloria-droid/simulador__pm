const ExcelJS = require('exceljs');
async function create() {
  const wb = new ExcelJS.Workbook(); wb.creator='Simulador PMM'; wb.created=new Date();
  const C={teal:'00838F',td:'004D40',tl:'B2DFDB',tv:'E0F7FA',w:'FFFFFF',b:'000000',
    red:'C62828',rl:'FFCDD2',gl:'C8E6C9',al:'FFF8E1',g:'ECEFF1',gd:'607D8B',amb:'E65100'};
  const fs=(c)=>({type:'pattern',pattern:'solid',fgColor:{argb:c}});
  const bt={top:{style:'thin',color:{argb:C.tl}},bottom:{style:'thin',color:{argb:C.tl}},
    left:{style:'thin',color:{argb:C.tl}},right:{style:'thin',color:{argb:C.tl}}};
  const bm={top:{style:'medium',color:{argb:C.teal}},bottom:{style:'medium',color:{argb:C.teal}},
    left:{style:'medium',color:{argb:C.teal}},right:{style:'medium',color:{argb:C.teal}}};
  const ac={horizontal:'center',vertical:'middle'},al={horizontal:'left',vertical:'middle'},
    ar={horizontal:'right',vertical:'middle'};
  const nE='#,##0 €', nP='0.0%';
  const fT={name:'Calibri',size:14,bold:true,color:{argb:C.w}};
  const fSub={name:'Calibri',size:10,italic:true,color:{argb:C.tl}};
  const fS={name:'Calibri',size:11,bold:true,color:{argb:C.w}};
  const fH={name:'Calibri',size:10,bold:true,color:{argb:C.w}};
  const fN={name:'Calibri',size:10,color:{argb:C.b}};
  const fB={name:'Calibri',size:10,bold:true,color:{argb:C.b}};
  const fBW={name:'Calibri',size:10,bold:true,color:{argb:C.w}};
  const fL={name:'Calibri',size:11,bold:true,color:{argb:C.td}};
  const fR={name:'Calibri',size:10,bold:true,color:{argb:C.red}};
  const fBig={name:'Calibri',size:14,bold:true,color:{argb:C.teal}};

  const prods=[{n:'Créditos',s:2000000},{n:'Descuento Comercial',s:8000000},
    {n:'Confirming',s:8000000},{n:'Comex',s:8000000},{n:'Préstamos',s:4000000}];
  const comps=['Empresa 1','Empresa 2','Empresa 3','Empresa 4'];
  const GL=8000000;
  const sd=[[500000,300000,200000,400000],[1000000,500000,800000,700000],
    [300000,200000,400000,100000],[0,100000,200000,0],[500000,300000,400000,300000]];

  const ws=wb.addWorksheet('Simulador PMM',{properties:{tabColor:{argb:C.teal}},
    views:[{showGridLines:false,state:'frozen',ySplit:6}]});
  [24,20,18,18,18,18,20,20,14].forEach((w,i)=>{ws.getColumn(i+1).width=w;});

  function sc(cell,o={}){
    if(o.f)cell.font=o.f;if(o.fl)cell.fill=o.fl;if(o.a)cell.alignment=o.a;
    if(o.b)cell.border=o.b;if(o.n)cell.numFmt=o.n;if(o.p)cell.protection=o.p;
  }
  const cl=(n)=>String.fromCharCode(64+n);
  const fR7=7,lR=11,fC=3,lC=6,gC=7,hC=8,iC=9,tR=12;
  // ALLSUM references all disposition cells directly (avoids circular ref with ERROR text)
  const ALLSUM='SUM($C$7:$F$11)';

  // === ROW 1: TITLE ===
  ws.mergeCells('A1:I1');
  sc(ws.getCell('A1'),{f:fT,fl:fs(C.teal),a:ac}); ws.getCell('A1').value='SIMULADOR PÓLIZA MULTIPRODUCTO MULTIEMPRESA (PMM)'; ws.getRow(1).height=38;

  // === ROW 2: SUBTITLE + DATE ===
  ws.mergeCells('A2:E2');
  sc(ws.getCell('A2'),{f:fSub,a:ac}); ws.getCell('A2').value='Contrato de Afianzamiento — Póliza Paraguas';
  ws.mergeCells('G2:I2');
  sc(ws.getCell('G2'),{f:{name:'Calibri',size:9,italic:true,color:{argb:C.gd}},a:ar});
  ws.getCell('G2').value={formula:'TEXT(TODAY(),"\"Fecha simulación: \"DD/MM/YYYY")'};
  ws.getRow(2).height=22;

  // === ROW 3: GLOBAL LIMIT + DISPONIBLE GLOBAL ===
  ws.mergeCells('A3:B3');
  sc(ws.getCell('A3'),{f:fL,a:ar}); ws.getCell('A3').value='LÍMITE GLOBAL PMM:';
  sc(ws.getCell('C3'),{f:fBig,fl:fs(C.tv),a:ac,b:bm,n:nE,p:{locked:false}});
  ws.getCell('C3').value=GL;
  // D3 spacer
  ws.mergeCells('E3:F3');
  sc(ws.getCell('E3'),{f:fL,a:ar}); ws.getCell('E3').value='DISPONIBLE GLOBAL:';
  ws.mergeCells('G3:H3');
  sc(ws.getCell('G3'),{f:fBig,fl:fs(C.gl),a:ac,b:bm,n:nE});
  ws.getCell('G3').value={formula:`MAX(0,$C$3-${ALLSUM})`};
  ws.getRow(3).height=32;

  // === ROW 5: SECTION HEADER ===
  ws.mergeCells('A5:I5');
  sc(ws.getCell('A5'),{f:fS,fl:fs(C.td),a:ac}); ws.getCell('A5').value='MATRIZ DE DISPOSICIONES POR EMPRESA Y PRODUCTO'; ws.getRow(5).height=28;

  // === ROW 6: TABLE HEADERS ===
  ['Tipo de Producto','Sublímite',...comps,'Total Dispuesto','Disponible','% Utilización'].forEach((h,i)=>{
    const c=ws.getRow(6).getCell(i+1); c.value=h; sc(c,{f:fH,fl:fs(C.teal),a:ac,b:bt});
  }); ws.getRow(6).height=28;

  // === ROWS 7-11: PRODUCTS ===
  prods.forEach((p,pi)=>{
    const r=fR7+pi, row=ws.getRow(r); row.height=24;
    const bg=pi%2===0?C.g:C.w;
    const RS=`SUM(C${r}:F${r})`;
    // A: name
    sc(row.getCell(1),{f:fB,fl:fs(bg),a:al,b:bt}); row.getCell(1).value=p.n;
    // B: sublimit (editable)
    sc(row.getCell(2),{f:fB,fl:fs(C.tv),a:ac,b:bt,n:nE,p:{locked:false}}); row.getCell(2).value=p.s;
    // C-F: dispositions (editable)
    comps.forEach((_,ci)=>{
      const c=fC+ci;
      sc(row.getCell(c),{f:fN,fl:fs(C.al),a:ac,b:bt,n:nE,p:{locked:false}});
      row.getCell(c).value=sd[pi][ci];
    });
    // G: Total Dispuesto — "ERROR" si excede sublímite O global
    sc(row.getCell(gC),{f:fB,fl:fs(C.tl),a:ac,b:bt,n:nE});
    row.getCell(gC).value={formula:`IF(OR(${RS}>B${r},${ALLSUM}>$C$3),"ERROR",${RS})`};
    // H: Disponible — suelo en 0, 0 si ERROR
    sc(row.getCell(hC),{f:fB,fl:fs(C.gl),a:ac,b:bt,n:nE});
    row.getCell(hC).value={formula:`IF(NOT(ISNUMBER(G${r})),0,MAX(0,MIN(B${r}-${RS},MAX(0,$C$3-${ALLSUM}))))`};
    // I: % Utilización
    sc(row.getCell(iC),{f:fB,fl:fs(bg),a:ac,b:bt,n:nP});
    row.getCell(iC).value={formula:`IF(NOT(ISNUMBER(G${r})),"—",IF(B${r}=0,0,${RS}/B${r}))`};
  });

  // === ROW 12: TOTAL ROW ===
  const tr=ws.getRow(tR); tr.height=28;
  sc(tr.getCell(1),{f:fBW,fl:fs(C.td),a:ac,b:bm}); tr.getCell(1).value='TOTAL PMM';
  sc(tr.getCell(2),{f:fBW,fl:fs(C.td),a:ac,b:bm,n:nE});
  tr.getCell(2).value={formula:`SUM(B${fR7}:B${lR})`};
  comps.forEach((_,ci)=>{
    const c=fC+ci,l=cl(c);
    sc(tr.getCell(c),{f:fBW,fl:fs(C.td),a:ac,b:bm,n:nE});
    tr.getCell(c).value={formula:`SUM(${l}${fR7}:${l}${lR})`};
  });
  // G12: Total global — ERROR si excede
  sc(tr.getCell(gC),{f:fBW,fl:fs(C.td),a:ac,b:bm,n:nE});
  tr.getCell(gC).value={formula:`IF(${ALLSUM}>$C$3,"ERROR",${ALLSUM})`};
  // H12: Disponible global
  sc(tr.getCell(hC),{f:fBW,fl:fs(C.td),a:ac,b:bm,n:nE});
  tr.getCell(hC).value={formula:`IF(NOT(ISNUMBER(G${tR})),0,MAX(0,$C$3-${ALLSUM}))`};
  // I12: % global
  sc(tr.getCell(iC),{f:fBW,fl:fs(C.td),a:ac,b:bm,n:nP});
  tr.getCell(iC).value={formula:`IF(NOT(ISNUMBER(G${tR})),"—",IF($C$3=0,0,${ALLSUM}/$C$3))`};

  // === CONDITIONAL FORMATTING ===
  // G: ERROR text → red
  ws.addConditionalFormatting({ref:`G${fR7}:G${tR}`,rules:[
    {type:'containsText',operator:'containsText',text:'ERROR',style:{fill:fs(C.rl),font:fR},priority:1}
  ]});
  // H: =0 → red
  ws.addConditionalFormatting({ref:`H${fR7}:H${tR}`,rules:[
    {type:'cellIs',operator:'equal',formulae:[0],style:{fill:fs(C.rl),font:fR},priority:2}
  ]});
  // I: >100% red, >80% amber
  ws.addConditionalFormatting({ref:`I${fR7}:I${lR}`,rules:[
    {type:'cellIs',operator:'greaterThan',formulae:[1],style:{fill:fs(C.rl),font:fR},priority:3},
    {type:'cellIs',operator:'greaterThan',formulae:[0.8],style:{fill:fs('FFF3E0'),font:{color:{argb:C.amb},bold:true}},priority:4}
  ]});
  // G3:H3 disponible global =0 → red
  ws.addConditionalFormatting({ref:'G3:H3',rules:[
    {type:'cellIs',operator:'lessThanOrEqual',formulae:[0],style:{fill:fs(C.rl),font:fR},priority:5}
  ]});

  // === ROW 14: INDICADORES GLOBALES ===
  ws.mergeCells('A14:E14');
  sc(ws.getCell('A14'),{f:fS,fl:fs(C.td),a:ac}); ws.getCell('A14').value='INDICADORES GLOBALES'; ws.getRow(14).height=28;
  ['Indicador','Importe','% sobre Límite','Estado',''].forEach((h,i)=>{
    const c=ws.getRow(15).getCell(i+1); c.value=h; sc(c,{f:fH,fl:fs(C.teal),a:ac,b:bt});
  });
  // Row 16: Límite
  ws.getRow(16).height=24;
  sc(ws.getCell('A16'),{f:fB,a:al,b:bt,fl:fs(C.g)}); ws.getCell('A16').value='Límite Global PMM';
  sc(ws.getCell('B16'),{f:fB,a:ac,b:bt,n:nE,fl:fs(C.g)}); ws.getCell('B16').value={formula:'$C$3'};
  sc(ws.getCell('C16'),{f:fB,a:ac,b:bt,n:nP,fl:fs(C.g)}); ws.getCell('C16').value=1;
  sc(ws.getCell('D16'),{f:fB,a:ac,b:bt,fl:fs(C.g)}); ws.getCell('D16').value='—';
  // Row 17: Dispuesto
  ws.getRow(17).height=24;
  sc(ws.getCell('A17'),{f:fB,a:al,b:bt}); ws.getCell('A17').value='Total Dispuesto';
  sc(ws.getCell('B17'),{f:fB,a:ac,b:bt,n:nE}); ws.getCell('B17').value={formula:`IF(${ALLSUM}>$C$3,"ERROR",${ALLSUM})`};
  sc(ws.getCell('C17'),{f:fB,a:ac,b:bt,n:nP}); ws.getCell('C17').value={formula:`IF(NOT(ISNUMBER(B17)),"—",IF($C$3=0,0,${ALLSUM}/$C$3))`};
  sc(ws.getCell('D17'),{f:fB,a:ac,b:bt}); ws.getCell('D17').value={formula:'IF(NOT(ISNUMBER(B17)),"⚠ EXCEDIDO","✓ OK")'};
  // Row 18: Disponible
  ws.getRow(18).height=24;
  sc(ws.getCell('A18'),{f:fB,a:al,b:bt,fl:fs(C.gl)}); ws.getCell('A18').value='Disponible Global';
  sc(ws.getCell('B18'),{f:fB,a:ac,b:bt,n:nE,fl:fs(C.gl)}); ws.getCell('B18').value={formula:`IF(NOT(ISNUMBER(B17)),0,MAX(0,$C$3-${ALLSUM}))`};
  sc(ws.getCell('C18'),{f:fB,a:ac,b:bt,n:nP,fl:fs(C.gl)}); ws.getCell('C18').value={formula:'IF($C$3=0,0,B18/$C$3)'};
  sc(ws.getCell('D18'),{f:fB,a:ac,b:bt,fl:fs(C.gl)}); ws.getCell('D18').value={formula:'IF(B18=0,"⚠ SIN DISPONIBLE","✓ DISPONIBLE")'};
  // Conditional: B17 ERROR → red, B18 =0 → red
  ws.addConditionalFormatting({ref:'B17',rules:[{type:'containsText',operator:'containsText',text:'ERROR',style:{fill:fs(C.rl),font:fR},priority:6}]});
  ws.addConditionalFormatting({ref:'B18',rules:[{type:'cellIs',operator:'equal',formulae:[0],style:{fill:fs(C.rl),font:fR},priority:7}]});
  ws.addConditionalFormatting({ref:'D17:D18',rules:[{type:'containsText',operator:'containsText',text:'EXCEDIDO',style:{fill:fs(C.rl),font:fR},priority:8},
    {type:'containsText',operator:'containsText',text:'SIN DISPONIBLE',style:{fill:fs(C.rl),font:fR},priority:9}]});

  // === ROW 20: INDICADORES POR PRODUCTO ===
  ws.mergeCells('A20:I20');
  sc(ws.getCell('A20'),{f:fS,fl:fs(C.td),a:ac}); ws.getCell('A20').value='INDICADORES POR PRODUCTO'; ws.getRow(20).height=28;
  ['Producto','Sublímite','Dispuesto','Disponible','% Sublímite','% Global','Estado Sublímite','Estado Global','Margen Global'].forEach((h,i)=>{
    const c=ws.getRow(21).getCell(i+1); c.value=h; sc(c,{f:fH,fl:fs(C.teal),a:ac,b:bt});
  });
  prods.forEach((p,pi)=>{
    const r=22+pi, dr=fR7+pi, row=ws.getRow(r); row.height=24;
    const bg=pi%2===0?C.g:C.w;
    const RS=`SUM(C${dr}:F${dr})`;
    sc(row.getCell(1),{f:fB,a:al,b:bt,fl:fs(bg)}); row.getCell(1).value=p.n;
    sc(row.getCell(2),{f:fN,a:ac,b:bt,n:nE,fl:fs(bg)}); row.getCell(2).value={formula:`B${dr}`};
    // C: Dispuesto (always numeric, uses SUM directly)
    sc(row.getCell(3),{f:fN,a:ac,b:bt,n:nE,fl:fs(bg)}); row.getCell(3).value={formula:RS};
    // D: Disponible (suelo 0)
    sc(row.getCell(4),{f:fB,a:ac,b:bt,n:nE,fl:fs(C.gl)}); row.getCell(4).value={formula:`MAX(0,H${dr})`};
    // E: % Sublímite
    sc(row.getCell(5),{f:fN,a:ac,b:bt,n:nP,fl:fs(bg)}); row.getCell(5).value={formula:`IF(B${r}=0,0,C${r}/B${r})`};
    // F: % Global
    sc(row.getCell(6),{f:fN,a:ac,b:bt,n:nP,fl:fs(bg)}); row.getCell(6).value={formula:`IF($C$3=0,0,C${r}/$C$3)`};
    // G: Estado Sublímite
    sc(row.getCell(7),{f:fB,a:ac,b:bt,fl:fs(bg)}); row.getCell(7).value={formula:`IF(C${r}>B${r},"⚠ EXCEDIDO",IF(B${r}=0,"—",IF(C${r}/B${r}>0.8,"⚡ ALTO","✓ OK")))`};
    // H: Estado Global
    sc(row.getCell(8),{f:fB,a:ac,b:bt,fl:fs(bg)}); row.getCell(8).value={formula:`IF(${ALLSUM}>$C$3,"⚠ EXCEDIDO","✓ OK")`};
    // I: Margen Global
    sc(row.getCell(9),{f:fN,a:ac,b:bt,n:nE,fl:fs(bg)}); row.getCell(9).value={formula:`MAX(0,$C$3-${ALLSUM})`};
  });
  // Cond fmt for indicators per product
  ws.addConditionalFormatting({ref:'D22:D26',rules:[{type:'cellIs',operator:'equal',formulae:[0],style:{fill:fs(C.rl),font:fR},priority:10}]});
  ws.addConditionalFormatting({ref:'G22:H26',rules:[{type:'containsText',operator:'containsText',text:'EXCEDIDO',style:{fill:fs(C.rl),font:fR},priority:11}]});
  ws.addConditionalFormatting({ref:'E22:E26',rules:[
    {type:'cellIs',operator:'greaterThan',formulae:[1],style:{fill:fs(C.rl),font:fR},priority:12},
    {type:'cellIs',operator:'greaterThan',formulae:[0.8],style:{fill:fs('FFF3E0'),font:{color:{argb:C.amb},bold:true}},priority:13}
  ]});

  // === LEYENDA (Row 28+) ===
  ws.mergeCells('A28:E28');
  sc(ws.getCell('A28'),{f:fS,fl:fs(C.td),a:ac}); ws.getCell('A28').value='LEYENDA'; ws.getRow(28).height=24;
  const legend=[
    ['🟡 Amarillo claro','Celda editable — disposiciones por empresa'],
    ['🔵 Azul claro','Celda editable — sublímite de producto / límite global'],
    ['🟢 Verde claro','Disponible (calculado automáticamente)'],
    ['🔴 Rojo / "ERROR"','Límite excedido — sublímite o global superado'],
    ['⚡ Ámbar','Utilización > 80% (alerta preventiva)'],
    ['✓ OK','Dentro de parámetros'],
  ];
  legend.forEach((l,i)=>{
    const r=29+i;
    sc(ws.getRow(r).getCell(1),{f:fB,a:al,b:bt}); ws.getRow(r).getCell(1).value=l[0];
    ws.mergeCells(`B${r}:E${r}`);
    sc(ws.getRow(r).getCell(2),{f:fN,a:al,b:bt}); ws.getRow(r).getCell(2).value=l[1];
  });

  // === DATA VALIDATION ===
  for(let r=fR7;r<=lR;r++){
    for(let c=fC;c<=lC;c++){
      ws.getCell(r,c).dataValidation={type:'decimal',operator:'greaterThanOrEqual',
        formulae:[0],showErrorMessage:true,errorTitle:'Valor no válido',error:'Las disposiciones no pueden ser negativas.'};
    }
    // Sublimit validation
    ws.getCell(r,2).dataValidation={type:'decimal',operator:'greaterThanOrEqual',
      formulae:[0],showErrorMessage:true,errorTitle:'Valor no válido',error:'El sublímite no puede ser negativo.'};
  }
  // Global limit validation
  ws.getCell('C3').dataValidation={type:'decimal',operator:'greaterThan',
    formulae:[0],showErrorMessage:true,errorTitle:'Valor no válido',error:'El límite global debe ser mayor que 0.'};

  // === NAMED RANGES ===
  wb.definedNames.addEx({name:'LimiteGlobal',refFormula:"'Simulador PMM'!$C$3"});
  wb.definedNames.addEx({name:'TotalDispuesto',refFormula:`'Simulador PMM'!$G$${tR}`});
  wb.definedNames.addEx({name:'DisponibleGlobal',refFormula:"'Simulador PMM'!$G$3"});

  // === PRINT ===
  ws.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0};

  // === INSTRUCTIONS SHEET ===
  const wh=wb.addWorksheet('Instrucciones',{properties:{tabColor:{argb:C.gd}},views:[{showGridLines:false}]});
  wh.getColumn(1).width=65; wh.getColumn(2).width=40;
  const ins=[
    ['INSTRUCCIONES DE USO — SIMULADOR PMM',''],['',''],
    ['1. LÍMITE GLOBAL (celda C3)',''],
    ['   Modifique el límite global de la póliza paraguas.',''],
    ['   El disponible global se muestra en G3:H3 y se actualiza automáticamente.',''],['',''],
    ['2. SUBLÍMITES POR PRODUCTO (columna B, filas 7-11)',''],
    ['   Modifique los sublímites de cada tipo de producto.',''],
    ['   Los sublímites pueden superar el límite global (son techos por producto).',''],['',''],
    ['3. DISPOSICIONES POR EMPRESA (celdas amarillas, C7:F11)',''],
    ['   Introduzca los importes dispuestos por cada empresa en cada producto.',''],
    ['   No se permiten valores negativos.',''],['',''],
    ['4. SISTEMA DE ALERTAS',''],
    ['   • "ERROR" (rojo): el total dispuesto supera el sublímite del producto',''],
    ['     O el total global supera el límite PMM.',''],
    ['   • Disponible = 0 (rojo): no queda margen disponible.',''],
    ['   • ⚡ ALTO (ámbar): utilización superior al 80%.',''],
    ['   • ✓ OK (verde): dentro de parámetros.',''],['',''],
    ['5. REGLA FUNDAMENTAL',''],
    ['   El límite global NO puede superarse. El disponible por producto',''],
    ['   se ajusta automáticamente respetando AMBOS techos (sublímite y global).',''],
    ['   Los valores nunca serán negativos: mínimo 0 o texto ERROR.',''],['',''],
    ['6. CELDAS EDITABLES',''],
    ['   • C3: Límite global PMM',''],
    ['   • B7:B11: Sublímites por producto',''],
    ['   • C7:F11: Disposiciones por empresa y producto',''],
    ['   Todas las demás celdas son de cálculo automático.',''],
  ];
  ins.forEach((r,i)=>{
    wh.getRow(i+1).getCell(1).value=r[0]; wh.getRow(i+1).getCell(2).value=r[1];
    if(i===0) wh.getRow(1).getCell(1).font={...fT,color:{argb:C.td}};
  });

  await wb.xlsx.writeFile('Simulador_PMM.xlsx');
  console.log('✅ Simulador_PMM.xlsx generado correctamente');
}
create().catch(console.error);
