const {SheetsMetadataDoc} = Import;

function fromId(id, sheetKey) {
  return SheetsMetadataDoc.fromId(id, sheetKey);
}

function getMetadata(doc, metadataKey=null, metadataValue=null) {
  return doc.getSpreadsheetMetadata({metadataKey, metadataValue});
}

function createMetadata(doc, metadataKey=null, metadataValue=null) {
  return doc.createSpreadsheetMetadata({metadataKey, metadataValue});
}

function updateMetadata(doc, metadataId, metadataKey, metadataValue, visibility) {
  return doc.updateSpreadsheetMetadata(metadataId, {metadataKey, metadataValue, visibility});
}
