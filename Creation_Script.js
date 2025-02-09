function myFunction() {
  // Create a new Google Form
  var form = FormApp.create('HEDA2024 - Automatically generated questionnaire');
  
  // Section 1: Identificación del Conjunto de Datos
  //Page1
  var page1 = form.addSectionHeaderItem().setTitle('1. IDENTIFICACION DEL CONJUNTO DE DATOS. (1)')
  .setHelpText('El conjunto de datos es una serie de datos estructurados, vinculados entre sí y agrupados dentro de una misma unidad temática y física, de forma que puedan ser procesados apropiadamente para obtener información.');
  
  var valEmailItem = form.addTextItem();
  valEmailItem.setTitle('Email').setRequired(true);
  
  var valDatasetNameItem = form.addTextItem();
  valDatasetNameItem.setTitle('NOMBRE DEL CONJUNTO DE DATOS').setRequired(true);
  
  var queInventoryNumberItem = form.addMultipleChoiceItem();
  queInventoryNumberItem.setTitle('¿Existe un número de inventario?');
  queInventoryNumberItem.setRequired(true);
  
  //Page2
  var page2 = form.addPageBreakItem().setTitle('1. IDENTIFICACION DEL CONJUNTO DE DATOS.  (1A)');
  var inventoryNumberTextItem = form.addTextItem();
  inventoryNumberTextItem.setTitle('Escriba el número de inventario').setRequired(true);
  
  //Page3
  var page3 = form.addPageBreakItem().setTitle('1. IDENTIFICACION DEL CONJUNTO DE DATOS. (2)');
  
  var queResponsibleInstitution = form.addMultipleChoiceItem();
  queResponsibleInstitution.setTitle('¿Existe una institución responsable del conjunto de datos?').setRequired(true);
  queResponsibleInstitution.setRequired(true);
  
  //Branching Page1
  queInventoryNumberItem.setChoices([
    queInventoryNumberItem.createChoice('No existe o no aplica',page3), 
    queInventoryNumberItem.createChoice('Se tiene el número de inventario',page2)
  ]);

  //Page4
  var page4 = form.addPageBreakItem().setTitle('1. IDENTIFICACION DEL CONJUNTO DE DATOS. (2A)');
  var institutionNameItem = form.addTextItem();
  institutionNameItem.setTitle('Escriba el nombre de la institución responsable del conjunto de datos').setRequired(true);
  
  //Page5
  var page5 = form.addPageBreakItem().setTitle('1. IDENTIFICACION DEL CONJUNTO DE DATOS. (3)'); 
  var queDatasetTypeItem = form.addMultipleChoiceItem();
  queDatasetTypeItem.setTitle('7. Seleccione el tipo del conjunto de datos.')
    .setChoiceValues(['Geográfico', 'Estadístico (encuesta, registro administrativo, censo, indicador)', 'Directorio', 'Padrón de beneficiarios', 'Other'])
    .setRequired(true);
  
  var datasetFormatItem = form.addCheckboxItem();
  datasetFormatItem.setTitle('8. ¿Cuál es el formato del conjunto de datos?')
    .setChoiceValues(['csv', 'xls', 'xlsx', 'Json', 'XML', 'rdf', 'shp', 'KML', 'WFS', 'WMS', 'txt', 'ODT', 'doc', 'docx', 'PDF', 'LOD', 'Other'])
    .setRequired(true);
  
  var generalCommentsItem = form.addParagraphTextItem();
  generalCommentsItem.setTitle('9. Comentarios generales sobre el conjunto de datos:');

  //Branching Page3
  queResponsibleInstitution.setChoices([
    queResponsibleInstitution.createChoice('No existe o no aplica', page5),
    queResponsibleInstitution.createChoice('Se tiene el número de inventario',page4)
  ]);
}