function codigoDeLaHojaDeCalculo(){
  return <Page ID> //Example "1fK3wISgnGsJramzAymZ06mZywOFIBBojz2sO7kZjHiY"
}

function codigoPrincipal() {
  var spa = SpreadsheetApp.openById(codigoDeLaHojaDeCalculo());
  var hojas = spa.getSheets();
  var dataSheetName = getMainSheetNames(hojas)
  var dataSheet = spa.getSheetByName(dataSheetName);
  var isPuntuacionMax = false; //Poner en true para remplazar la puntuacion aplicable con la maxima.
  var tituloDimension = ["Demanda.","Apertura de Datos.","Calidad.","Total"];
  var subTitulosDemanda = ["Relevancia Del conjunto de datos.", "Prioridades de gobierno.", "Grupos vulnerables.","Planificación y orientación de recursos.", "Identificación de areas de ahorro.",
  "Colaboración entre las diferentes instituciones gubernamentales.", "Seguridad ciudadana.", "Experiencia y satisfacción ciudadana.", "Factores ambientales.", 
  "Nuevas areas de actividad y crecimiento económico.","Planificación económica del territorio."];
  var board = [];
  worksheet = spa.getSheetByName("Resultados")
    if(worksheet){   
       spa.deleteSheet(worksheet);
    }
  worksheet = spa.insertSheet("Resultados");

  for(var linea=2; linea<=dataSheet.getLastRow(); linea++){
  //for(var linea=2; linea<3; linea++){ //for testing purposes.
    var localBoard = [];
    
    //Creacion de la Dimension de Informacion.
    var infoRange = dataSheet.getRange(linea,1,linea,11);
    var infoArray = infoRange.getValues()[0];
    
    var dim1 = [];
    var columnaDim1=12;
    for(var subdimension = 0; subdimension < 11; subdimension++){
      var auxiliar = subdimension == 1 ? 10 : 4;
      infoRange = dataSheet.getRange(linea,columnaDim1,1,auxiliar);
      if(subdimension == 0){
        dim1[subdimension]= valoresDim1A(infoRange.getValues(), subTitulosDemanda[subdimension]);
      }else if(subdimension == 1){
        dim1[subdimension]= valoresDim1B(infoRange.getValues(), subTitulosDemanda[subdimension]);
      }else{
        dim1[subdimension]= valoresDim1C(infoRange.getValues(), subTitulosDemanda[subdimension]);
      }
      columnaDim1+=auxiliar;
    }

    var dim2 = [];
    infoRange = dataSheet.getRange(linea,62,1,5); //de BJ a BN Licencia Abierta.
    dim2[0] = valoresDA1(infoRange.getValues(), "Licencia abierta");
    //Creacion de la Dimension 3 Datos Abiertos - Diseño -Metadatos.
    infoRange = dataSheet.getRange(linea,67,1,4);//de BO a BR Metadatos.
    dim2[1] = valoresDA2(infoRange.getValues(), "Metadatos"); 
    //Accesibilidad.
    infoRange = dataSheet.getRange(linea,71,1,8);// de BS a BZ Accesibilidad
    dim2[2] = valoresDA3(infoRange.getValues(), "Accesibilidad");

    var dim3 = [];
    infoRange = dataSheet.getRange(linea,79,1,6); //de CA a CF Comparabilidad.
    dim3[0] = valoresC1(infoRange.getValues(), "Comparabilidad.");
    infoRange = dataSheet.getRange(linea,85,1,11); //de CG a CP Metadatos.
    dim3[1] = valoresC2(infoRange.getValues(), "Metadatos (Formato)");
    infoRange = dataSheet.getRange(linea,95,1,12); //de CQ a DB Metadatos.
    dim3[2] = valoresC3(infoRange.getValues(), "Metadatos (Registros)");
    infoRange = dataSheet.getRange(linea,107,1,14); //de DC a DP Metadatos.
    dim3[3] = valoresC4(infoRange.getValues(), "Consistencia de los Datos");
    infoRange = dataSheet.getRange(linea,121,1,6); //de DQ a DV Metadatos.
    dim3[4] = valoresC5(infoRange.getValues(), "Oportunidad y Puntualidad");


    var diseArray = infoRange.getValues()[0];
    //Creacion de la Dimension 4 Comparabilidad.
    infoRange = dataSheet.getRange(linea,79,linea,127);
    var compaArray = infoRange.getValues()[0]; 
    //Definicion de la hoja de la encuesta
    
    var fecha = (""+infoArray[0]);
    var gmt = fecha.indexOf("GMT");
    gmt = gmt ? gmt : fecha.length-1;
    fecha = fecha.substring(0,gmt)
    localBoard[0] = fecha;

    //Logger.log("fecha " + fecha);
    var worksheet = spa.getSheetByName(""+fecha)
    if(worksheet){   
       spa.deleteSheet(worksheet);
    }
    worksheet = spa.insertSheet(fecha);

    //Presentation de la informacion
    var tituloCD =''
    var celdaActiva = worksheet.getRange("A1:L1");
    if(infoArray[3].length >0){
      celdaActiva = confCeldaTitle(celdaActiva);
      tituloCD = "Conjunto de datos: "+infoArray[3]
      celdaActiva.setValue(tituloCD)
      localBoard[1] = infoArray[3];
    }else{
      tituloCD = "Conjunto de datos sin nombre."
      celdaActiva.setValue(tituloCD)
      localBoard[1] = tituloCD;
    }
    
    
    celdaActiva = worksheet.getRange("A2:L2");
    celdaActiva = confCeldaRef(celdaActiva);
    if(infoArray[7].length >0){
      celdaActiva.setValue("Elaborado por: "+infoArray[1]+" para el "+infoArray[7])
    }else{
      celdaActiva.setValue("Elaborado por: "+infoArray[1])
    }

    //linea de inicio
    var ldi = 28;
    //columna de inicio
    var  cdi = 3;
    //Dimension 2
    celdaActiva = worksheet.getRange(columnToLetter(cdi)+ ldi +":"+columnToLetter(cdi+7)+ldi);
    celdaActiva.merge();
    celdaActiva = confCeldaTitle(celdaActiva);
    celdaActiva.setValue("Detalles");
    ldi++
    celdaActiva = worksheet.getRange(columnToLetter(cdi)+ ldi +":"+columnToLetter(cdi+7)+ldi);
    celdaActiva.merge();
    celdaActiva = confCeldaSubTitle(celdaActiva);
    celdaActiva.setValue("Dimensión 1: Demanda");
    ldi++;
    celdaActiva = worksheet.getRange(columnToLetter(cdi) + ldi);
    celdaActiva.setValue("Sub-dimensión");
    celdaActiva = worksheet.getRange(columnToLetter(cdi) + ldi+":"+columnToLetter(cdi+3) + ldi);
    celdaActiva.merge();
    celdaActiva = worksheet.getRange(columnToLetter(cdi+3) + ldi);
    celdaActiva.setValue("Puntuación");
    celdaActiva = worksheet.getRange(columnToLetter(cdi+4) + ldi);
    celdaActiva.setValue("Pun. Aplicable");
    celdaActiva = worksheet.getRange(columnToLetter(cdi+5) + ldi);
    celdaActiva.setValue("Pun. Max");
    celdaActiva = worksheet.getRange(columnToLetter(cdi+7) + ldi);
    celdaActiva.setValue("Evidencias");
    celdaActiva = worksheet.getRange(columnToLetter(cdi+6) + ldi);
    celdaActiva.setValue("Evaluación");
    celdaActiva = worksheet.getRange(columnToLetter(cdi) + ldi+":"+columnToLetter(cdi+7) + ldi)
    celdaActiva = confCeldaTituloColumna(celdaActiva);
    //celdaActiva.setHorizontalAlignment("center");
    ldi++;
    celdaActiva = worksheet.getRange(columnToLetter(cdi) + (ldi-2)+":"+columnToLetter(cdi+7) + (ldi+11))
    celdaActiva.setBorder(true, true, true, true, false, false);



    var lineaActiva = 6;
    //worksheet = fillDimensionSoloMax(dim1,worksheet,ldi, cdi, true, "Utilidad e Impacto");
    worksheet = fillDimension(dim1,worksheet,ldi, cdi, true, "Utilidad e Impacto", isPuntuacionMax);

    ldi+=12;
    //Dim2 Datos Abiertos
    celdaActiva = worksheet.getRange(columnToLetter(cdi)+ ldi +":"+columnToLetter(cdi+7)+ldi);
    celdaActiva.merge();
    celdaActiva = confCeldaSubTitle(celdaActiva);
    celdaActiva.setValue("Dimensión 2: Apertura de Datos");
    ldi++;
    celdaActiva = worksheet.getRange(columnToLetter(cdi) + ldi);
    celdaActiva.setValue("Sub-dimensión");
    celdaActiva = worksheet.getRange(columnToLetter(cdi) + ldi+":"+columnToLetter(cdi+3) + ldi);
    celdaActiva.merge();
    celdaActiva = worksheet.getRange(columnToLetter(cdi+3) + ldi);
    celdaActiva.setValue("Puntuación");
    celdaActiva = worksheet.getRange(columnToLetter(cdi+4) + ldi);
    celdaActiva.setValue("Pun. Aplicable");
    celdaActiva = worksheet.getRange(columnToLetter(cdi+5) + ldi);
    celdaActiva.setValue("Pun. Max");
    celdaActiva = worksheet.getRange(columnToLetter(cdi+7) + ldi);
    celdaActiva.setValue("Evidencias");
    celdaActiva = worksheet.getRange(columnToLetter(cdi+6) + ldi);
    celdaActiva.setHorizontalAlignment("center");
    celdaActiva.setValue("Evaluación");
    celdaActiva = worksheet.getRange(columnToLetter(cdi) + ldi+":"+columnToLetter(cdi+7) + ldi)
    celdaActiva = confCeldaTituloColumna(celdaActiva);
    ldi++
    celdaActiva = worksheet.getRange(columnToLetter(cdi) + (ldi-2)+":"+columnToLetter(cdi+7) + (ldi+4))
    celdaActiva.setBorder(true, true, true, true, false, false);



    celdaActiva = worksheet.getRange(columnToLetter(cdi)+ ldi +":"+columnToLetter(cdi+7)+ldi);
    celdaActiva.merge();
    celdaActiva = confCeldaHighlight(celdaActiva);
    celdaActiva.setHorizontalAlignment("center");
    celdaActiva.setValue("Diseño:");
    ldi++;
    //lineaActiva = 19;
    //worksheet = fillDimensionSoloMax(dim2,worksheet,ldi, cdi,true, "Accesibilidad");
    worksheet = fillDimension(dim2,worksheet,ldi, cdi,true, "Accesibilidad", isPuntuacionMax);
    ldi+=4;
    //Calidad
    celdaActiva = worksheet.getRange(columnToLetter(cdi)+ ldi +":"+columnToLetter(cdi+7)+ldi);
    celdaActiva.merge();
    celdaActiva = confCeldaSubTitle(celdaActiva);
    celdaActiva.setValue("Dimensión 3: Calidad");
    ldi++;
    celdaActiva = worksheet.getRange(columnToLetter(cdi) + ldi);
    celdaActiva.setValue("Sub-dimensión");
    celdaActiva = worksheet.getRange(columnToLetter(cdi) + ldi+":"+columnToLetter(cdi+3) + ldi);
    celdaActiva.merge();
    celdaActiva = worksheet.getRange(columnToLetter(cdi+3) + ldi);
    celdaActiva.setValue("Puntuación");
    //Por si cambian de opinion
    celdaActiva = worksheet.getRange(columnToLetter(cdi+4) + ldi);
    celdaActiva.setValue("Pun. Aplicable");
    celdaActiva = worksheet.getRange(columnToLetter(cdi+5) + ldi);
    celdaActiva.setValue("Pun. Max");
    celdaActiva = worksheet.getRange(columnToLetter(cdi+7) + ldi);
    celdaActiva.merge();
    celdaActiva.setValue("Evidencias");
    celdaActiva = worksheet.getRange(columnToLetter(cdi+6) + ldi);
    celdaActiva.setHorizontalAlignment("center");
    celdaActiva.setValue("Evaluación");
    celdaActiva = worksheet.getRange(columnToLetter(cdi) + ldi+":"+columnToLetter(cdi+7) + ldi)
    celdaActiva = confCeldaTituloColumna(celdaActiva);
    
    ldi++;
    celdaActiva = worksheet.getRange(columnToLetter(cdi) + (ldi-2)+":"+columnToLetter(cdi+7) + (ldi+4))
    celdaActiva.setBorder(true, true, true, true, false, false);

    //lineaActiva = 25;
    //worksheet = fillDimensionSoloMax(dim3,worksheet,ldi, cdi, false, "");
    worksheet = fillDimension(dim3,worksheet,ldi, cdi, false, "", isPuntuacionMax);

    var resumen = [];
    resumen[0] = valoresDimension(dim1);
    resumen[1] = valoresDimension(dim2);
    resumen[2] = valoresDimension(dim3);

    var aceptable = [];
    //Aceptabilidad de la Demanda
    if(resumen[0][0] > 0){
      //Logger.log(dim1[1]);
      if(dim1[1][0] > 4){
        if(dim1[2][0] > 0 || dim1[3][0] > 0 || dim1[4][0] > 0 || dim1[5][0] > 0 || dim1[6][0] > 0 || dim1[7][0] > 0 || dim1[8][0] > 0 || dim1[9][0] > 0 || dim1[10][0] > 0){
          aceptable[0] = "Aceptable";
        }else{
          aceptable[0] = "No aceptable, Utilidad e Impacto = 0";
        }
      }else{
        aceptable[0] = "No aceptable, Prioridades < 5";
      }
    }else{
      aceptable[0] = "No aceptable, Demanda = 0";
    }
    aceptable[0]= aceptable[0] ? aceptable[0]: "Error";

    //Aceptabilidad de los Datos Abiertos
    if(resumen[1][0] > 0){
      if(dim2[0][0] > 0){
        if(dim2[2][0] > 0){
          aceptable[1] = "Aceptable";
        }else{
          aceptable[1] = "No aceptable, Acceso a los datos = 0";
        }
      }else{
        aceptable[1] = "No aceptable, Licencia Abierta = 0";
      }
    }else{
      aceptable[1] = "No aceptable, Datos Abiertos = 0";
    }
    aceptable[1]= aceptable[1] ? aceptable[1]: "Error";

    //Aceptabilidad de la calidad
    if(resumen[2][0] > 0){
          aceptable[2] = "Aceptable";
    }else{
      aceptable[2] = "No aceptable, Calidad < 3"
    }
    aceptable[2]= aceptable[2] ? aceptable[2]: "Error";

    celdaActiva = worksheet.getRange("B4:K4");
    celdaActiva.merge();
    celdaActiva = confCeldaSubTitle(celdaActiva);
    celdaActiva.setFontColor("#003E19");
    celdaActiva.setValue("Estadísticas por Dimensión");
    celdaActiva = worksheet.getRange("E5:F5");
    celdaActiva = confCeldaObtenido(celdaActiva);
    celdaActiva.setValue("Puntuación \nObtenida");
    if(isPuntuacionMax){
      celdaActiva = worksheet.getRange("G5:H5");
      celdaActiva = confCeldaMax(celdaActiva);
      celdaActiva.setValue("Puntuación \nMaxima");
    }else{
      celdaActiva = worksheet.getRange("G5:H5");
      celdaActiva = confCeldaApp(celdaActiva);
      celdaActiva.setValue("Puntuación \nAplicable");
    }
    celdaActiva = worksheet.getRange("I5:K5");
    celdaActiva.merge();
    celdaActiva.setValue("Aceptabilidad de acuerdo al Manual HEDA");
    celdaActiva = confCeldaObtenido(celdaActiva);
    
    //TODO Reducir
    celdaActiva = worksheet.getRange("B6:D6");
    celdaActiva = confCeldaHighlight(celdaActiva);
    celdaActiva.setValue("Demanda.");
    celdaActiva.setHorizontalAlignment("right");
    worksheet = createResultTableCase(worksheet, "E",6,resumen[0], isPuntuacionMax);
    celdaActiva = worksheet.getRange("I6:K6");
    celdaActiva.merge();
    celdaActiva.setValue(aceptable[0]);
    if(isNegation(aceptable[0])){
      celdaActiva = confCeldaNotOk(celdaActiva);
    }else{
      celdaActiva = confCeldaOk(celdaActiva);
    }

    celdaActiva = worksheet.getRange("B7:D7");
    celdaActiva.setValue("Apertura de Datos.");
    celdaActiva.setHorizontalAlignment("right");
    celdaActiva = confCeldaHighlight(celdaActiva);
    worksheet = createResultTableCase(worksheet, "E",7,resumen[1], isPuntuacionMax);
    celdaActiva = worksheet.getRange("I7:K7");
    celdaActiva.merge();
    celdaActiva.setValue(aceptable[1]);
    if(isNegation(aceptable[1])){
      celdaActiva = confCeldaNotOk(celdaActiva);
    }else{
      celdaActiva = confCeldaOk(celdaActiva);
    }

    celdaActiva = worksheet.getRange("B8:D8");
    celdaActiva.setValue("Calidad.");
    celdaActiva.setHorizontalAlignment("right");
    celdaActiva = confCeldaHighlight(celdaActiva);
    worksheet = createResultTableCase(worksheet, "E",8,resumen[2], isPuntuacionMax);
    celdaActiva = worksheet.getRange("I8:K8");
    celdaActiva.merge();
    celdaActiva.setValue(aceptable[2]);
    if(isNegation(aceptable[2])){
      celdaActiva = confCeldaNotOk(celdaActiva);
    }else{
      celdaActiva = confCeldaOk(celdaActiva);
    }

    var total = generarTotal(resumen);
    celdaActiva = worksheet.getRange("B9:D9");
    celdaActiva.setValue("Resultado Total.");
    celdaActiva.setFontWeight("bold");
    celdaActiva.setHorizontalAlignment("right");
    celdaActiva = confCeldaHighlight(celdaActiva);

    worksheet = createResultTableCase(worksheet, "E",9,total);
    celdaActiva = worksheet.getRange("B9:H9");
    celdaActiva.setBorder(true, true, true, true, false, false);
    celdaActiva = worksheet.getRange("E9:J9");
    celdaActiva.setFontWeight("bold");
    celdaActiva.setFontStyle("italic");

    cdi = letterToColumn("E");
    ldi = 13;
    celdaActiva = worksheet.getRange(columnToLetter(cdi+1)+ldi);
    celdaActiva = confHidden(celdaActiva);
    celdaActiva.setValue("Puntuación \nObtenida");

    //todo reducir
    celdaActiva = worksheet.getRange(columnToLetter(cdi)+ldi);
    celdaActiva = confHidden(celdaActiva);
    celdaActiva.setValue("Demanda.");
//    worksheet = createReferenceTable(worksheet, columnToLetter(cdi+1),ldi,resumen[0]);
    worksheet = createReferenceTableCase(worksheet, columnToLetter(cdi+1),ldi,resumen[0],isPuntuacionMax);
    ldi++;
    celdaActiva = worksheet.getRange(columnToLetter(cdi)+ldi);
    celdaActiva.setValue("Apertura de Datos.");
    celdaActiva = confHidden(celdaActiva);
//    worksheet = createReferenceTable(worksheet, columnToLetter(cdi+1),ldi,resumen[1]);
    worksheet = createReferenceTableCase(worksheet, columnToLetter(cdi+1),ldi,resumen[1],isPuntuacionMax);
    ldi++;
    celdaActiva = worksheet.getRange(columnToLetter(cdi)+ldi);
    celdaActiva.setValue("Calidad.");
    celdaActiva.setHorizontalAlignment("right");
    celdaActiva = confHidden(celdaActiva);
//    worksheet = createReferenceTable(worksheet, columnToLetter(cdi+1),ldi,resumen[2]);
    worksheet = createReferenceTableCase(worksheet, columnToLetter(cdi+1),ldi,resumen[2],isPuntuacionMax);

    var tableRange = worksheet.getRange("E13:F16");

    var chart = newBarChart(worksheet,tableRange,"Porcentaje de Puntuación por Dimensión",10,1,5,5);
    worksheet.insertChart(chart);
    
    localBoard[2] = resumen;
    localBoard[3] = total;
    localBoard[4] = aceptable;
    board[linea-2] = localBoard;


  }

  worksheet = spa.getSheetByName("Resultados")

  var ldi = 1;
  var cdi = 1;
  var resBoard = generarEstadisticasDelBoard(board);
  var acpBoard = contarAceptaciones(board);

  var celdaActiva = worksheet.getRange(columnToLetter(cdi)+ ldi +":"+columnToLetter(cdi+14)+ldi);
  celdaActiva = confCeldaTitle(celdaActiva)
  celdaActiva.setValue("Resultados de los Conjuntos de Datos.");
  ldi ++;
  var celdaActiva = worksheet.getRange(columnToLetter(cdi+2)+ ldi +":"+columnToLetter(cdi+12)+ldi);
  celdaActiva = confCeldaTituloColumna(celdaActiva)
  celdaActiva.setValue("El siguiente es el resultado de los " + board.length + " Conjuntos de Datos analizados");
  ldi ++;
  celdaActiva = worksheet.getRange(columnToLetter(cdi+2)+ ldi +":"+columnToLetter(cdi+5)+ldi);
  celdaActiva = confCeldaHighlight(celdaActiva);
  celdaActiva.setValue("Dimensión");
  celdaActiva = worksheet.getRange(columnToLetter(cdi+6)+ ldi +":"+columnToLetter(cdi+7)+ldi);
  celdaActiva = confCeldaHighlight(celdaActiva);
  celdaActiva.setValue("Puntuación");
  celdaActiva = worksheet.getRange(columnToLetter(cdi+8)+ ldi +":"+columnToLetter(cdi+9)+ldi);
  celdaActiva = confCeldaHighlight(celdaActiva);
  celdaActiva.setValue("Puntuación Aplicable");
  celdaActiva = worksheet.getRange(columnToLetter(cdi+10)+ ldi +":"+columnToLetter(cdi+12)+ldi);
  celdaActiva = confCeldaHighlight(celdaActiva);
  celdaActiva.setValue("Aceptables de acuerdo al Manual HEDA");
ldi++;

  for(var k =0; k<resBoard.length; k++){
    celdaActiva = worksheet.getRange(columnToLetter(cdi+2)+ ldi +":"+columnToLetter(cdi+5)+ldi);
    celdaActiva = confCeldaHighlight(celdaActiva)
    celdaActiva.setHorizontalAlignment("left");
    celdaActiva.setValue(tituloDimension[k]);
    celdaActiva = worksheet.getRange(columnToLetter(cdi+6)+ ldi +":"+columnToLetter(cdi+7)+ldi);
    celdaActiva.merge();
    celdaActiva.setValue(isPuntuacionMax ? resBoard[k][0]/resBoard[k][2] : resBoard[k][0]/resBoard[k][1]);
    celdaActiva.setNumberFormat('0.00%');
    celdaActiva.setHorizontalAlignment("right");
    celdaActiva = worksheet.getRange(columnToLetter(cdi+8)+ ldi +":"+columnToLetter(cdi+9)+ldi);
    celdaActiva.merge();
    celdaActiva.setValue(isPuntuacionMax ? resBoard[k][0]+"/"+resBoard[k][2] : resBoard[k][0]+"/"+resBoard[k][1]);
    celdaActiva.setHorizontalAlignment("left");
    celdaActiva = worksheet.getRange(columnToLetter(cdi+10)+ ldi +":"+columnToLetter(cdi+11)+ldi);
    celdaActiva.merge();
    celdaActiva.setValue(acpBoard[k]);
    ldi++;
  }
  ldi++;
 

  celdaActiva = worksheet.getRange(columnToLetter(cdi)+ ldi +":"+columnToLetter(cdi+14)+ldi);
  celdaActiva = confCeldaTitle(celdaActiva);
  celdaActiva.setValue("Detalles");
  ldi++;
  celdaActiva = worksheet.getRange(columnToLetter(cdi)+ldi+":"+columnToLetter(cdi+7)+ldi);
  celdaActiva = confCeldaTituloColumna(celdaActiva);
  celdaActiva.setValue("Conjunto de Datos");
  celdaActiva = worksheet.getRange(columnToLetter(cdi+8)+ldi+":"+columnToLetter(cdi+11)+ldi);
  celdaActiva = confCeldaTituloColumna(celdaActiva);
  celdaActiva.setValue("Puntuación");
  celdaActiva = worksheet.getRange(columnToLetter(cdi+12)+ldi+":"+columnToLetter(cdi+14)+ldi);
  celdaActiva = confCeldaTituloColumna(celdaActiva);
  celdaActiva.setValue("Aceptabilidad de acuerdo al Manual HEDA");
  ldi++;
  celdaActiva = worksheet.getRange(columnToLetter(cdi)+ldi+":"+columnToLetter(cdi+2)+ldi);
  celdaActiva = confCeldaHighlight(celdaActiva);
  celdaActiva.setValue("Fecha y hora");
  celdaActiva = worksheet.getRange(columnToLetter(cdi+3)+ldi+":"+columnToLetter(cdi+7)+ldi);
  celdaActiva = confCeldaHighlight(celdaActiva);
  celdaActiva.setValue("Nombre del Conjunto de Datos");
  celdaActiva = worksheet.getRange(columnToLetter(cdi+8)+ldi+":"+columnToLetter(cdi+9)+ldi);
  celdaActiva = confCeldaHighlight(celdaActiva);
  celdaActiva.setValue("Obtenida");
  celdaActiva = worksheet.getRange(columnToLetter(cdi+10)+ldi+":"+columnToLetter(cdi+11)+ldi);
  if(isPuntuacionMax){
    celdaActiva.setValue("Máxima");
    celdaActiva = confCeldaObtenido(celdaActiva);
  }else{
    celdaActiva.setValue("Aplicable");
    celdaActiva = confCeldaObtenido(celdaActiva);
  }
  celdaActiva = worksheet.getRange(columnToLetter(cdi+12)+ldi);
  celdaActiva = confCeldaHighlight(celdaActiva);
  celdaActiva.setValue("Demanda");
  celdaActiva = worksheet.getRange(columnToLetter(cdi+13)+ldi);
  celdaActiva = confCeldaHighlight(celdaActiva);
  celdaActiva.setValue("AD");
  celdaActiva = worksheet.getRange(columnToLetter(cdi+14)+ldi);
  celdaActiva = confCeldaHighlight(celdaActiva);
  celdaActiva.setValue("Calidad");
  ldi++;
  worksheet = fillBoard(worksheet,ldi,board, isPuntuacionMax);
  worksheet.getRange("A1").activate();
}

function contarAceptaciones(array){
  var result = [];
  result[0] = 0;
  result[1] = 0;
  result[2] = 0;
  for (var i = 0; i<array.length; i++){
    var aux = array[i];
    aux = aux[4];
    for (var j = 0; j<aux.length; j++){
      result [j] = result[j] + (isNegation(aux[j])?0:1);
    }
  }
  return result;
}

function generarEstadisticasDelBoard(array){
  var result = [];
  for (var i = 0; i<array.length; i++){
    result[i] =[];
    var aux = array[i];
    var aux2 = aux[2]
    for(var k = 0; k<3; k++){
      var aux3 = aux2[k];
      result[i][k] = [];
      result[i][k][0] = 0;
      result[i][k][1] = 0;
      result[i][k][2] = 0;
      for (var j = 0; j<aux3.length; j++){
        result[i][k][j] = result[i][k][j] + parseInt(aux3[j]);
      }
    }
  }
  
  var realResult =[];
  realResult[0]=[];
  realResult[1]=[];
  realResult[2]=[];
  for (var i = 0; i<result.length; i++){
    var k=0;
    for(k=0;k<3;k++){
      for(var j=0; j<3;j++){
        realResult[k][j] = realResult[k][j] ? realResult[k][j] + result[i][k][j] : result[i][k][j];
      }
    }
  }
  realResult[3] = [];
  for (var i = 0; i<array.length; i++){
    var aux = array[i];
    for (var j = 0; j<aux.length; j++){
      realResult[3][j] = realResult[3][j] ? realResult[3][j] +  parseInt(aux[3][j]) : parseInt(aux[3][j]);
    }
  }
  

  return realResult;
}


function generarTotal(datos){
  var result = [];
  result[0] = 0
  result[1] = 0
  result[2] = 0
  for (var i = 0; i < datos.length; i++){
    var dato = datos[i];
    for (var j = 0; j < dato.length; j++){
      result[j] = result[j] + dato[j];
    }
  }
  return result;
}

function letterToColumn(letter) {
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 65 + 1) * Math.pow(26, length - i - 1);
  }
  return column;
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function fillBoard(sheet,y,board, bool){
  var switchy = (bool?2:1);
  var x =1;
  var i;
  var celda;
  for(i=0; i<board.length; i++){
    celda = sheet.getRange((columnToLetter(x))+(y+i)+":"+(columnToLetter(x+2))+(y+i));
    celda.merge();
    celda.setValue(board[i][0]);
    celda = sheet.getRange(columnToLetter(x+3)+(y+i)+":"+(columnToLetter(x+7))+(y+i));
    celda.merge();
    celda.setValue(board[i][1]);
    celda = sheet.getRange(columnToLetter(x+8)+(y+i)+":"+(columnToLetter(x+9))+(y+i));
    celda.merge();
    celda.setValue(board[i][3][0]/board[i][3][switchy]);
    celda.setNumberFormat('0.00%');
    celda = sheet.getRange(columnToLetter(x+10)+(y+i)+":"+(columnToLetter(x+11))+(y+i));
    celda.merge();
    celda.setValue(board[i][3][0]+"/"+board[i][3][switchy]);
    celda.merge();
    celda = sheet.getRange(columnToLetter(x+12)+(y+i));
    celda.merge();
    if(isNegation(board[i][4][0])){
      celda.setValue("No aceptable");
      celda = confCeldaNotOk(celda);
    }else{
      celda.setValue(board[i][4][0]);
      celda = confCeldaOk(celda)
    }
    celda.setHorizontalAlignment("center");
    celda = sheet.getRange(columnToLetter(x+13)+(y+i));
    celda.merge();
    if(isNegation(board[i][4][1])){
      celda.setValue("No aceptable");
      celda = confCeldaNotOk(celda);
    }else{
      celda.setValue(board[i][4][1]);
      celda = confCeldaOk(celda)
    }
    celda.setHorizontalAlignment("center");
    celda = sheet.getRange(columnToLetter(x+14)+(y+i));
    celda.merge();
    if(isNegation(board[i][4][2])){
      celda.setValue("No aceptable");
      celda = confCeldaNotOk(celda);
    }else{
      celda.setValue(board[i][4][2]);
      celda = confCeldaOk(celda)
    }
    celda.setHorizontalAlignment("center");
  }
  celda = sheet.getRange((columnToLetter(x))+(y-2)+":"+(columnToLetter(x+7))+(y+(i-1)));
  celda.setBorder(true, true, true, true, false, false);
  celda = sheet.getRange((columnToLetter(x+8))+(y-2)+":"+(columnToLetter(x+11))+(y+(i-1)));
  celda.setBorder(true, true, true, true, false, false);
  celda = sheet.getRange((columnToLetter(x+12))+(y-2)+":"+(columnToLetter(x+14))+(y+(i-1)));
  celda.setBorder(true, true, true, true, false, false);

  return sheet;
}

function createResultTableCase(sheet, x, y, array, bool){
  var switchy = (bool ? 2:1);
  var valorPorcentual = array[0]/array[switchy]
  var celda = sheet.getRange(x+(y)+":"+columnToLetter(letterToColumn(x)+1)+(y));
  celda.merge();
  celda.setValue(valorPorcentual);
  celda.setNumberFormat('0.00%');
  x = columnToLetter(letterToColumn(x)+2);

  var celda = sheet.getRange(x+(y)+":"+columnToLetter(letterToColumn(x)+1)+(y));
  celda.merge();
  celda.setValue(array[0]+"/"+array[switchy]);
  return sheet;
}


function createReferenceTable(sheet, x, y, array){
  for (var i = 0; i <array.length; i++){
    var celda = sheet.getRange(x+(y))
    celda = confHidden(celda);
    celda.setValue(array[i]);
    celda.setNumberFormat('0.00%');
    x = columnToLetter(letterToColumn(x)+1);
  }
  return sheet;
}

function createReferenceTableCase(sheet, x, y, array, bool){
  var switchy = (bool ? 2:1);
  var celda = sheet.getRange(x+(y))
  celda = confHidden(celda);
  celda.setValue(array[0]/array[switchy]);
  celda.setNumberFormat('0.00%');
  x = columnToLetter(letterToColumn(x)+1);
  return sheet;
}

function getMainSheetNames(hojas){
  var patron = /^[A-Za-z]{3}\s[A-Za-z]{3}\s\d{2}\s\d{4}\s\d{2}\:\d{2}\:\d{2}/;
  for (var i=0; i<hojas.length;i++){
    var auxiliar = hojas[i].getName();
    if(auxiliar != "Resultados"){
      if(!patron.test(auxiliar)){
        return auxiliar;
      }
    }
  }
  Logger.log("no encontrado")
  return hojas[0].getName();
}

function isInArray(fecha, hojas){
  for (indice = 0; indice < hojas.length; indice++){
    if(hojas[indice].getNamel === fecha){
      return true;
    }
  }
  return false;
}

function newPieChart(sheet, range, title, p1, p2, l, h){
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(range))
    .setOption('title', title)
    .setPosition(p1,p2,l,h) 
    .build();
  return chart;
}

function newBarChart(sheet, range, title, p1, p2, l, h){
  var chart;
  chart = sheet.newChart()
  //.asBarChart()
  .setChartType(Charts.ChartType.BAR)
  .addRange(range)
  //.setOption('legend', 'none')
  .setPosition(p1,p2,l,h) 
  .setOption('hAxis.title', 'Porcentaje de Evaluación')
  .setOption('vAxis;title', 'Dimensión')
  .setOption('width', 1200)
  //.setOption('chartArea.left', l)
  .setOption('series', {0: {color: '#2C0EAE'}})
  .setOption('series', {0: {labelInLegend : 'Puntuación'}})
  .setOption('series', {0: {dataLabel: 'value'}})
  .setOption('title', title)
  .setOption('titleTextStyle', {fontSize: 20, bold: true, color: "Black",alignment: 'center'})
  //.setOption('legend', {'position':'top'})
  .build();
  return chart;
  
}

function valoresDim1A(values, texto){
  //The score values of the 115 questions.
  //Demanda
  var valores = values[0];
  //Logger.log(valores);
  var result = [];
  result[0] = 0;
  result[1] = 0;
  result[2] = 0;
  result[3] = texto;
  //Logger.log(valores[0]);
  if(valores[0] !== '' && valores[0]== "Si se sabe."){
    result[0] = result[0]+ 1;
    result[1] = result[1]+1;
    result[0] = result[0] + (valores[1] !== '' && valores[1]== "Las evidencias son correctamente documentadas." ? 2 : 0);
    //Logger.log(valores[1]);
    result[1] = result[1]+2;
    result[0] = result[0] + (valores[2] !== '' && valores[2]== "Retroalimentación y además se conoce la satisfacción de los usuarios." ? 2 : 0);
    result[1] = result[1]+2;
    result[1] = result[1]+1;//documento
    //prioridades de gobierno
    if(valores[3] && valores[3] !== ''){
      result[0] = result[0] + 1;
      result[4] = valores[3];
    }
  }else{
    result[4] ="No aplica.";
  }
  result[2]=1+2+2+1;
  return result;
}

function valoresDim1B(values, texto){
  var valores = values[0];
  var result = [];
  result[0] = 0;
  result[1] = 0;
  result[2] = 0;
  result[3] = texto;
  if(isSi(valores[0])){
    result[0] = result[0] + (valores[1]!== '' && isSi(valores[1]) ? 3 : 0);
    result[1] = result[1] + 3;
    result[0] = result[0] + (valores[2]!== '' && valores[2]== "Existe un marco legal o reglamento definido que regula el Conjunto de Datos." ? 2 : 0);
    result[1] = result[1] + 2;
    result[0] = result[0] + (valores[3]!== '' && isSi(valores[3]) ? 5 : 0);
    result[1] = result[1] + 5;
    result[0] = result[0] + (valores[4]!== '' && isSi(valores[4]) ? 2 : 0);
    result[1] = result[1] + 2;
    result[0] = result[0] + (valores[5]!== '' && isSi(valores[5]) ? 2 : 0);
    result[1] = result[1] + 2;
    result[0] = result[0] + (valores[6]!== '' && isSi(valores[6]) ? 2 : 0);
    result[1] = result[1] + 2;
    result[0] = result[0] + (valores[7]!== '' && valores[7]== "Multiples procesos." ? 1 : 0);
    result[1] = result[1] + 1;
    result[0] = result[0] + (valores[8]!== '' && valores[8]== "Existen evidencias de que se ha utilizado el conjunto de datos para la evaluación de impacto de programas públicos." ? 2 : 0);
    result[1] = result[1] + 2;
    result[1] = result[1]+1;//documento
    if(valores[3] && valores[3] !== ''){
      result[0] = result[0] + 1;
      result[4]=valores[9];
    }
  }else{
    result[4] ="No aplica.";
  }
  result[2] = 3 + 2 + 5 + 2 + 2 + 2 + 1 + 2 + 1;
  return result;
}

function valoresDim1C(values, texto){
  var valores = values[0];
  var result = [];
  result[0] = 0;
  result[1] = 0;
  result[2] = 0;
  result[3] = texto;
  if(isSi(valores[0])){
    result[0] = result[0] + 1;
    result[1] = result[1]+1;
    result[0] = result[0] + (isSi(valores[1]) ? 2 : 0);
    result[1] = result[1]+2;
    result[0] = result[0] + (isSi(valores[2]) ? 2 : 0);
    result[1] = result[1]+2;
    result[1] = result[1]+1;//documento
    if(valores[3] && valores[3] !== ''){
      result[0] = result[0] + 1;
      result[4] = valores[3];
    }
  }else{
    result[4] ="No aplica.";
  }
  result[2] = 1 + 2 + 2 + 1;
  return result;
}

function valoresDA1(values, texto){
  var valores = values[0];
  var result = [];
  //Logger.log(valores);
  result[0] = 0;
  result[1] = 0;
  result[2] = 0;
  result[3] = texto;
  //var values= []; // [1]Puntuacion Optenida [0]Valor de la pregunta
  result[0] = result[0] + (isSi(valores[0]) ? 1 : 0);
  result[1] = result[1]+1;
  result[0] = result[0] + (isSi(valores[1]) ? 5 : 0);
  result[1] = result[1]+5;
  result[0] = result[0] + (valores[2] ? valores[2] : 0);
  result[1] = result[1]+5;
  result[0] = result[0] + (isSi(valores[3]) ? 3 : 0);
  result[1] = result[1]+3;
  result[0] = result[0] + (isNegation(valores[4]) ? 2 : 0);
  result[1] = result[1]+2;
  result[2] = 1 + 5 + 5 + 3 + 2;
  return result;
}

function valoresDA2(values, texto){
  var valores = values[0];
  var result = [];
  //Logger.log(valores);
  result[0] = 0;
  result[1] = 0;
  result[2] = 0;
  result[3] = texto;
  if(isSi(valores[0])){
    result[0] = result[0] + 5;
    result[1] = result[1]+5;
    result[0] = result[0] + (isSi(valores[1]) ? 5 : 0);
    result[1] = result[1]+5;
    result[0] = result[0] + (isSi(valores[2]) ? 5 : 0);
    result[1] = result[1]+5;
    result[0] = result[0] + (isSi(valores[3]) ? 5 : 0);
    result[1] = result[1]+5;
  }else{
    result[4] ="No aplica.";
  }
  result[2] = 5 + 5 + 5 + 5;
  return result;
}

function valoresDA3(values, texto){
  var valores = values[0];
  var result = [];
  //Logger.log(valores);
  result[0] = 0;
  result[1] = 0;
  result[2] = 0;
  result[3] = texto;
  //var values= []; // [1]Puntuacion Optenida [0]Valor de la pregunta
  result[0] = result[0] + (isNegation(valores[1]) ? 5 : 0);
  result[1] = result[1]+5;
  result[0] = result[0] + (isNegation(valores[2]) ? 5 : 0);
  result[1] = result[1]+5;
  result[0] = result[0]+ (valores[3] !== '' && valores[3]== "Un unico pago." ? 1 : 0);
  result[1] = result[1]+1;
  result[0] = result[0]+ (isSi(valores[4]) ? 2 : 0);
  result[1] = result[1]+2;
  result[0] = result[0]+ (isSi(valores[5]) ? 1 : 0);
  result[1] = result[1]+1;
  result[0] = result[0]+ (isSi(valores[6]) ? 1 : 0);
  result[1] = result[1]+1;
  result[1] = result[1]+3;
  if(valores[7] && valores[7] !== ''){
    result[0] = result[0] + 3;
    result[4] = valores[3];
  }
  result[2] = result[1];
  return result
}

//Calidad
//Comparabilidad
function valoresC1(values, texto){
  var valores = values[0];
  var result = [];
  //Logger.log(valores);
  result[0] = 0;
  result[1] = 0;
  result[2] = 0;
  result[3] = texto;
  result[0] = result[0] + (valores[1] === "" ? 0 : (isNegation(valores[0]) ? 0 : 2));
  result[1] = result[1] + 2;
  result[0] = result[0] + (isSi(valores[1]) ? 3 : 0);
  result[1] = result[1] + 3;
  result[0] = result[0] + (isSi(valores[2]) ? 1 : 0);
  result[1] = result[1] + 1;
  result[0] = result[0] + (isSi(valores[3]) ? 1 : 0);
  result[1] = result[1] + 1;
  result[0] = result[0] + (isSi(valores[4]) ? 1 : 0);
  result[1] = result[1] + 1;
  result[4] = '';
  if (valores[4] && valores[5].length > 5){
    result[0] = result[0] +3;
    result[4] = "Clasificadores usados:\n" + valores[5];
  }
  result[1] = result[1]+3;
  result[2] = result[1];
  return result;
}

//Metadatos
function valoresC2(values, texto){
  var valores = values[0];
  var result = [];
  //Logger.log(valores);
  result[0] = 0;
  result[1] = 0;
  result[2] = 0;
  result[3] = texto;
  result[4] = ''
  result[0] = result[0] + (isSi(valores[0]) ? 1 : 0);
  result[1] = result[1] + 1;
  result[0] = result[0] + (isSi(valores[1]) ? 1 : 0);
  result[1] = result[1] + 1;
  result[0] = result[0] + (valores[2] === "" ? 0 : ((isSi(valores[2]) ? 3 : 0)));
  result[1] = result[1] + 3;
  if (valores[3] && valores[3].length > 2){
    result[0] = result[0] + 5;
    result[3] = "Formatos usados:\n" + valores[3];
  }
  result[1] = result[1] + 5;
  result[0] = result[0] + (isNegation(valores[4]) ? 0 : 5);
  result[1] = result[1] + 5;
  result[0] = result[0]+ (valores[5] !== '' && valores[5]== "Existe un sub-registro considerable." ? 1 : 0);
  result[1] = result[1] + 1;
  result[1] = result[1] + 1;
  if (valores[6] && valores[6].length > 2){
    result[0] = result[0] + 1;
    result[4] = "\nSub-registro, casos particulares:\n" + valores[6];
  }
  result[0] = result[0] + (isNegation(valores[7]) ? 0 : 5);
  result[1] = result[1] + 5;
  result[0] = result[0]+ (valores[8] !== '' && valores[8]== "Existe un sobre-registro considerable." ? 1 : 0);
  result[1] = result[1] + 1;
  if (valores[9] && valores[9].length > 2){
    result[0] = result[0] + 1;
    result[4] = "\nSobre-registro, casos particulares:\n" + valores[10];
  }
  result[2] = 1+1+3+5+5+1+1+5+1+1;
  return result;
}

//registros
function valoresC3(values, texto){
  var valores = values[0];
  var result = [];
  //Logger.log(valores);
  result[0] = 0;
  result[1] = 0;
  result[2] = 0;
  result[3] = texto;
  result[4] = ''
  result[0] = result[0] + (isSi(valores[0]) ? 1 : 0);
  result[1] = result[1] + 1;
  result[0] = result[0] + (valores[1] === "" ? 0 : ((isSi(valores[1]) ? 0 : 4)));
  result[1] = result[1] + 4;
  result[0] = result[0] + (isSi(valores[2]) ? 1 : 0);
  result[1] = result[1] + 1;
  result[0] = result[0] + (valores[3] === "" ? 0 : ((isSi(valores[3]) ? 0 : 4)));
  result[1] = result[1] + 4;
  result[0] = result[0] + (isSi(valores[4]) ? 1 : 0);
  result[1] = result[1] + 1;
  result[0] = result[0] + (valores[5] === "" ? 0 : ((isSi(valores[5]) ? 0 : 4)));
  result[1] = result[1] + 4;
  result[0] = result[0] + (valores[6] === "" ? 0 : ((isNegation(valores[6]) ? 0 : 5)));
  result[1] = result[1] + 5;
  result[0] = result[0] + (isSi(valores[7]) ? 1 : 0);
  result[1] = result[1] + 1;
  result[0] = result[0] + (valores[8] === "" ? 0 : ((isSi(valores[8]) ? 0 : 4)));
  result[1] = result[1] + 4;
  result[0] = result[0] + (isSi(valores[9]) ? 1 : 0);
  result[1] = result[1] + 1;
  result[0] = result[0] + (valores[10] === "" ? 0 : ((isSi(valores[10]) ? 0 : 4)));
  result[1] = result[1] + 4;
  result[0] = result[0] + (valores[0] === "" ? 0 : ((isNegation(valores[11]) ? 0 : 2)));
  result[1] = result[1] + 2;
  result[2] = result[2] + 1 + 4 + 1 + 4 + 1 + 4 + 5 + 1 + 4 + 1 + 4 + 2;
  return result;
}

function valoresC4(values, texto){
  var valores = values[0];
  var result = [];
  //Logger.log(valores);
  result[0] = 0;
  result[1] = 0;
  result[2] = 0;
  result[3] = texto;
  result[4] = ''
  result[0] = result[0] + (isSi(valores[0]) ? 2 : 0);
  result[1] = result[1] + 2;
  result[0] = result[0] + (valores[1] === "" ? 0 : ((isSi(valores[1]) ? 0 : 3)));
  result[1] = result[1] + 3;
  result[0] = result[0] + (isSi(valores[2]) ? 1 : 0);
  result[1] = result[1] + 1;
  result[0] = result[0] + (isSi(valores[3]) ? 3 : 0);
  result[1] = result[1] + 3;
  if(isSi(valores[4])){
    result[0] = result[0] + 1;  
    result[4] = "Formatos usados:\nJson o RDF"
  }else{
    result[4] = "Formatos usados:\n" + valores[4];
  }
  result[1] = result[1] + 1;
  result[0] = result[0] + (isSi(valores[5]) ? 3 : 0);
  result[1] = result[1] + 3;
  result[0] = result[0] + (isSi(valores[6]) ? 3 : 0);
  result[1] = result[1] + 3;
  result[0] = result[0] + (isSi(valores[7]) ? 3 : 0);
  result[1] = result[1] + 3;
  if(isSi(valores[8])){
    result[0] = result[0] + 2;
    result[4] = "\nDDI y/o NTM"
  }else{
    result[4] = "\n" + valores[8];
  }
  result[1] = result[1] + 2;
  result[0] = result[0] + (isSi(valores[9]) ? 1 : 0);
  result[1] = result[1] + 1;
  result[0] = result[0] + (isSi(valores[10]) ? 2 : 0);
  result[1] = result[1] + 2;
  
  result[0] = result[0] + (isSi(valores[12]) ? 1 : 0);
  result[1] = result[1] + 1;
  if (valores[13] && valores[13] == "El reporte de calidad usa el formato DDI y / o NTM"){
    result[0] = result[0] + 2;
  }else{
    result[4] = "\n" + valores[13];
  }
  result[1] = result[1] + 2;
  if(valores[11] && valores[11] !== ''){
    result[0] = result[0] + 2;
    result[4] = valores[11];
  }
  result[1] = result[1] + 2;
  result[2] = 2+3+1+3+1+3+3+3+2+1+2+1+2+2;
  return result;
}

function valoresC5(values, texto){
  var valores = values[0];
  var result = [];
  //Logger.log(valores);
  result[0] = 0;
  result[1] = 0;
  result[2] = 0;
  result[3] = texto;
  result[4] = ''
  if(isSi(valores[0])){
    result[0] = result[0] + 3;
    result[0] = result[0] + (valores[1] !== '' && valores[1]== "El retraso fue significativo." ? 0 : 1);
    result[0] = result[0] + (isSi(valores[2]) ? 1: 0);
    result[0] = result[0] + valores[3] == "" ? 0 : ((isNegation(valores[3]) ? 0: 4));
    result[0] = result[0] + (valores[4] !== '' && valores[4]== "Mas del 90%." ? 5 : 0);
    result[4] = "El Conjunto de Datos ha sido publicado."
    result[1] = 3+1+1+4+5;
  }else{
    var addedQuestionString = valores[5].substring(0,1);
    if(valores[5] || valores[6]){
      if(addedQuestionString && addedQuestionString == "M"){
        result[0] = result[0] + ((valores[5] !== '' && valores[5]== "Mas del 90%.") ? 5 : 0);
      }else if (isSi(valores[5])){
        result[0] = result[0] + 1;
      }
      result[0] = (valores[6] && isSi(valores[6])) ? result[0] + 1: result[0];
    }
    result[4] = isSi(valores[5]) ? "El Conjunto de Datos es candidato a la apertura." : "El Conjunto de Datos NO es candidato a la apertura.";
    result[1] = 5+1;
  }
  result[2] = result[2] + 3 + 1 + 1 + 4 + 5; 
  return result;
}


function isSi(cadena){
  //Logger.log(cadena);
  //return cadena && cadena.length > 2 ? (cadena.substring(0,2) == 'Si' || aux2 == "Sí") : false;
  if(cadena && cadena.length > 2){
    var aux = cadena.substring(0,2);
    if (aux == 'Si'|| aux =="Sí" || aux =="sí" || aux =="si"){
      return true;
    }
    return false;
  }
}

function isNegation(cadena){
  //Logger.log(cadena);
  if(cadena && cadena.length > 3){
    var aux = cadena.indexOf("no ");
    aux = aux == -1 ? cadena.indexOf("no,"):aux;
    aux = aux == -1 ? cadena.indexOf("no."):aux;
    aux = aux == -1 ? cadena.indexOf("No "):aux;
    aux = aux == -1 ? cadena.indexOf("No,"):aux;
    aux = aux == -1 ? cadena.indexOf("No."):aux;
    return (aux  !== -1);
  }else{
    if(cadena && cadena.length > 2){
      if (aux == 'No'|| aux =="no"){
        return true;
      }else{
        return false;
      }
    }
  }
}

function confCeldaTitle(celda){
  celda.setBackground("#2C0EAE")
  celda.setHorizontalAlignment("center");
  celda.setFontSize(16);
  celda.merge();
  celda.setFontColor("WHITE");
  return celda;
}
function confCeldaRef(celda){
  celda.setBackground("#8F9145")
  celda.setHorizontalAlignment("center");
  celda.setFontSize(12);
  celda.merge();
  celda.setFontColor("WHITE");
  return celda;
}

function confCeldaSubTitle(celda){
  celda.setFontWeight("bold");
  celda.setHorizontalAlignment("center");
  celda.setFontSize(14);
  celda.setFontColor("#2C0EAE");
  return celda;
}

function confCeldaHighlight(celda){
  celda.setFontWeight("bold");
  celda.setHorizontalAlignment("center");
  celda.merge();
  celda.setFontSize(11);
  return celda;
}

function confHidden(celda){
  celda.setBackground("WHITE");
  celda.setFontColor("WHITE");
  return celda;
}

function confCeldaObtenido(celda){
  //celda.setBackground("#2C0EAE");
  //celda.setFontColor("WHITE");
  celda.setFontWeight("bold");
  celda.setHorizontalAlignment("center");
  celda.merge();
  return celda;
}

function confCeldaMax(celda){
  //celda.setBackground("#262625")
  //celda.setFontColor("WHITE");
  celda.setFontWeight("bold");
  celda.setHorizontalAlignment("center");
  celda.merge();
  return celda;
}

function confCeldaApp(celda){
  //celda.setBackground("#AE0606")
  //celda.setFontColor("WHITE");
  celda.setFontWeight("bold");
  celda.setHorizontalAlignment("center");
  celda.merge();
  return celda;
}

function confNoAplica(celda){
  celda.setFontColor("#8A0003");
  celda.setHorizontalAlignment("center");
  celda.setFontWeight("italic");
  return celda;
}

function confCeldaNotOk(celda){
  celda.setBackground("#AE0606");
  celda.setFontColor("WHITE");
  celda.setHorizontalAlignment("center");
  celda.setFontWeight("italic");
  return celda;
}

function confCeldaAcceptable(celda){
  celda.setBackground("YELLOW");
  celda.setFontColor("BLACK");
  celda.setHorizontalAlignment("center");
  celda.setFontWeight("italic");
  return celda;
}

function confCeldaOk(celda){
  celda.setBackground("#006600");
  celda.setFontColor("WHITE");
  celda.setHorizontalAlignment("center");
  celda.setFontWeight("italic");
  return celda;
}

function confCeldaTituloColumna(celda){
  celda.setBackground("TEAL");
  celda.setFontColor("WHITE");
  celda.setHorizontalAlignment("center");
  celda.setFontWeight("bold");
  celda.setFontStyle("italic");
  celda.merge();
  return celda;
}


function valoresDimension(dim){
  var outputs = [];
  outputs[0] = 0;
  outputs[1] = 0;
  outputs[2] = 0;
  for (var i =0; i<dim.length; i++){
    outputs[0] = outputs[0] + dim[i][0];
    outputs[1] = outputs[1] + dim[i][1];
    outputs[2] = outputs[2] + dim[i][2];
  }
  return outputs;
}

function fillDimension(dim, sheet, line, column, isAdd, texto, bool){
  var switchy = (bool?2:1);
  for (var i =0; i<dim.length; i++){
    var celda = sheet.getRange(line,column);
    celda.setValue(dim[i][3]);
    if(dim[i].length > 3 && dim[i][4] === "No aplica."){
      celda = sheet.getRange(line,column+3,1,5);  
      celda.merge();
      celda = confNoAplica(celda);
      celda.setValue("No Aplica.")
    }else{
      if(isAdd && i === 2){
        celda = sheet.getRange(line,column,1,8);
        celda.merge();
        celda.setHorizontalAlignment("center");
        celda = confCeldaHighlight(celda);
        celda.setValue(texto);
        line = line + 1;
      }
      celda = sheet.getRange(line,column);
      celda.setValue(dim[i][3]);
      celda = sheet.getRange(line,column+3);
      celda.setValue(dim[i][0]);
      celda = sheet.getRange(line,column+4);
      celda.setValue(dim[i][1]);
      celda = sheet.getRange(line,column+5);
      celda.setValue(dim[i][2]);
      celda = sheet.getRange(line,column+7);
      celda.setValue(dim[i].length > 4 ? dim[i][4] : '');
      celda = sheet.getRange(line,column+6);
      celda.setValue(dim[i][0]/dim[i][switchy]);
      celda.setNumberFormat('0.00%');
      celda.setHorizontalAlignment("center");
    }
    line = line + 1
  }
  return sheet;
}

function fillDimensionSoloMax(dim, sheet, line, column, isAdd, texto){
  for (var i =0; i<dim.length; i++){
    var celda = sheet.getRange(line,column);
    celda.merge();
    celda.setValue(dim[i][3]);
    if(dim[i].length > 3 && dim[i][4] === "No aplica."){
      celda = sheet.getRange(line,column+3,1,5);  
      celda.merge();
      celda = confNoAplica(celda);
      celda.setValue("No Aplica.")
    }else{
      if(isAdd && i === 2){
        celda = sheet.getRange(line,column,1,7);
        celda.merge();
        celda = confCeldaHighlight(celda);
        celda.setValue(texto);
        line = line + 1;
      }
      celda = sheet.getRange(line,column+3);
      celda.setValue(dim[i][0]);
      celda = sheet.getRange(line,column+4);
      celda.setValue(dim[i][2]);
      celda = sheet.getRange(line,column+6);
      celda.setValue(dim[i].length > 4 ? dim[i][4] : '');
      celda = sheet.getRange(line,column+5);
      celda.setValue(dim[i][0]/dim[i][2]);
      celda.setNumberFormat('0.00%');
      celda.setHorizontalAlignment("center");
    }
    line = line + 1
  }
  return sheet;
}