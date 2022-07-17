var estilos_sheet = PropertiesService.getDocumentProperties();

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Menú principal')
    .addItem('Mostrar barra lateral','mostrarBarraLateral')
    .addToUi();
}

function mostrarBarraLateral(){
  var ui = HtmlService.createTemplateFromFile('BarraLateral')
  .evaluate()
  .setTitle('Barra lateral propia');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function guardarEstilo(numEstilo){

  // borramos previamente los estilos
  eliminarEstilo(numEstilo);

  // obtenemos la celda activa
  var celda = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();

  // guardamos los bordes
  guardarBordes(celda, numEstilo);
 
  estilos_sheet.setProperty('colorLetra'+numEstilo, celda.getFontColor())
              .setProperty('colorFondo'+numEstilo, celda.getBackground())
              .setProperty('sizeFuente'+numEstilo, celda.getFontSize()+'');

  return{ colorFondo: estilos_sheet.getProperty('colorFondo'+numEstilo),
          colorLetra: estilos_sheet.getProperty('colorLetra'+numEstilo)};
}

function guardarBordes(celda, numEstilo){
  var bordes = celda.getBorder();

  if(bordes!=null){
    var bordeTop = bordes.getTop();
    var bordeLeft = bordes.getLeft();
    var bordeRight = bordes.getRight();
    var bordeBottom = bordes.getBottom();

    if(bordeTop.getColor() !=null && bordeTop.getBorderStyle() != null){
      estilos_sheet.setProperty('bordeTopCo'+numEstilo, bordeTop.getColor().asRgbColor().asHexString())
                    .setProperty('bordeTopSt'+numEstilo, bordeTop.getBorderStyle());
    }
    if(bordeLeft.getColor() !=null && bordeLeft.getBorderStyle() != null){
      estilos_sheet.setProperty('bordeLefCo'+numEstilo, bordeLeft.getColor().asRgbColor().asHexString())
                    .setProperty('bordeLefSt'+numEstilo, bordeLeft.getBorderStyle());
    }
    if(bordeRight.getColor() !=null && bordeRight.getBorderStyle() != null){
      estilos_sheet.setProperty('bordeRigCo'+numEstilo, bordeRight.getColor().asRgbColor().asHexString())
                    .setProperty('bordeRigSt'+numEstilo, bordeRight.getBorderStyle());
    }
    if(bordeBottom.getColor() !=null && bordeBottom.getBorderStyle() != null){
      estilos_sheet.setProperty('bordeBotCo'+numEstilo, bordeBottom.getColor().asRgbColor().asHexString())
                    .setProperty('bordeBotSt'+numEstilo, bordeBottom.getBorderStyle());
    }
  }
}


function aplicarEstilo(numEstilo){

  borrarEstilos();

  var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdas = hojaActual.getActiveRange();

  celdas.setFontColor(estilos_sheet.getProperty('colorLetra'+numEstilo))
    .setBackground(estilos_sheet.getProperty('colorFondo'+numEstilo))
    .setFontSize(estilos_sheet.getProperty('sizeFuente'+numEstilo))
    .setValue('Estilo'+numEstilo);

  // Aplicamos bordes
  if(comprobarBordes('Top', numEstilo)){
    celdas.setBorder(true, null, null, null, null, null, estilos_sheet.getProperty('bordeTopCo'+numEstilo), obtenerEnumBorde(estilos_sheet.getProperty('bordeTopSt'+numEstilo)))
  }
  if(comprobarBordes('Left', numEstilo)){
    celdas.setBorder(null, true, null, null, null, null, estilos_sheet.getProperty('bordeLefCo'+numEstilo), obtenerEnumBorde(estilos_sheet.getProperty('bordeLefSt'+numEstilo)))
  }
  if(comprobarBordes('Bottom', numEstilo)){
    celdas.setBorder(null, null, true, null, null, null, estilos_sheet.getProperty('bordeBotCo'+numEstilo), obtenerEnumBorde(estilos_sheet.getProperty('bordeBotSt'+numEstilo)))
  }
  if(comprobarBordes('Right', numEstilo)){
    celdas.setBorder(null, null, null, true, null, null, estilos_sheet.getProperty('bordeRigCo'+numEstilo), obtenerEnumBorde(estilos_sheet.getProperty('bordeRigSt'+numEstilo)))
  }
}

function comprobarBordes(borde, numEstilo){
  switch(borde){
    case 'Top': return estilos_sheet.getProperty('bordeTopCo'+numEstilo) != null;
    case 'Left': return estilos_sheet.getProperty('bordeTLefCo'+numEstilo) != null;
    case 'Bottom': return estilos_sheet.getProperty('bordeBotCo'+numEstilo) != null;
    case 'Right': return estilos_sheet.getProperty('bordeRigCo'+numEstilo) != null;
  }
}

function obtenerEnumBorde(tipoBorde){
  switch(tipoBorde){
    case 'DOTTED': return SpreadsheetApp.BorderStyle.DOTTED;
    case 'DASHED': return SpreadsheetApp.BorderStyle.DASHED;
    case 'SOLID': return SpreadsheetApp.BorderStyle.SOLID;
    case 'SOLID_THICK': return SpreadsheetApp.BorderStyle.SOLID_THICK;
    case 'DOUBLE': return SpreadsheetApp.BorderStyle.DOUBLE;
    default: return null;
  }
}

function eliminarEstilo(estilo){
  estilos_sheet.deleteProperty('colorLetra'+estilo);
  estilos_sheet.deleteProperty('colorFondo'+estilo);
  estilos_sheet.deleteProperty('sizeFuente'+estilo);

  // bordes
  estilos_sheet.deleteProperty('bordeTopCo'+estilo);
  estilos_sheet.deleteProperty('bordeTopSt'+estilo);
  estilos_sheet.deleteProperty('bordeRigCo'+estilo);
  estilos_sheet.deleteProperty('bordeRigSt'+estilo);
  estilos_sheet.deleteProperty('bordeLefCo'+estilo);
  estilos_sheet.deleteProperty('bordeLefSt'+estilo);
  estilos_sheet.deleteProperty('bordeBotCo'+estilo);
  estilos_sheet.deleteProperty('bordeBotSt'+estilo);

}

function cargarEstilos(){
  return estilos_sheet.getProperties();
}

function borrarEstilos(){
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear({formatOnly: true});
}

function borrarTodo(){
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear();
}