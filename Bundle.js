const Import = Object.create(null);
(function (exports) {

  const isEmpty = (obj) => Object.keys(obj).length === 0;
  const getMatched = (json) => json.matchedDeveloperMetadata.pop().developerMetadata;

  class Search {
    /**
     * returns sheet id by sheetKey
     * if not present, return null
     */
    constructor (endpoints) {
      this.endpoints = endpoints;
    }

    setRows (sheetId, howMany) { 
      const addRowsRequest = this.endpoints.spreadsheets.batchUpdate();
      addRowsRequest.updateSheetProperties(sheetId, {
        gridProperties: {
          rowCount: howMany
        }
      }, 'gridProperties.rowCount');
      const response = addRowsRequest.fetch();
      const json = response.json;
      if (!response.ok) throw new Error("Issue when adding rows");
    }

    setColumns (sheetId, howMany) { 
      const addColumnsRequest = this.endpoints.spreadsheets.batchUpdate();
      addColumnsRequest.updateSheetProperties(sheetId, {
        gridProperties: {
          columnCount: howMany
        }
      }, 'gridProperties.columnCount');
      const response = addColumnsRequest.fetch();
      if (!response.ok) throw new Error("Issue when adding columns");
    }

    /**
     * Returns the sheetId stored in md, or create it if not present
     * TODO: Add check if tab already exists when creating
     */
    getSheetId(sheetKey) {
      const metadataKey = `ctt_sheetKey_${sheetKey}`;
      const request = this.endpoints.developerMetadata.search();
      request.bySpreadsheetLoc({metadataKey});
      const json = request.fetch().json;
      if (isEmpty(json)) {
        // create new tab, get the id, save it in metadata, store in this.sheet
        const addSheetReq = this.endpoints.spreadsheets.batchUpdate();
        addSheetReq.addSheet(sheetKey, {
          gridProperties: {
            rowCount: 2,
            columnCount: 1,
            frozenRowCount: 1
          }
        });
        const addSheetResp = addSheetReq.fetch();
        if (!addSheetResp.ok) throw new Error(addSheetResp.json.error.message);
        const reply = addSheetResp.json.replies.pop();
        // got it:
        const sheetId = reply.addSheet.properties.sheetId;

        // now create sheet md
        const addSheetMdReq = this.endpoints.spreadsheets.batchUpdate();
        addSheetMdReq.createMetaData({
          metadataKey, 
          metadataValue: sheetId.toString()
        }, {
          location: {
            spreadsheet: true
          },
          visibility: "DOCUMENT"
        });
        const result = addSheetMdReq.fetch();  // should work
        return sheetId;
      };
      return parseInt(getMatched(json).metadataValue);
    }

    getLastColumn(sheetId) {
      const request = this.endpoints.spreadsheets.get();
      const json = request.fetch().json;
      const sheet = json.sheets.filter(sh => sh.properties.sheetId === sheetId).pop();
      return sheet.properties.gridProperties.columnCount;
    }

    getLastRow(sheetId) {
      const request = this.endpoints.spreadsheets.get();
      const json = request.fetch().json;
      const sheet = json.sheets.filter(sh => sh.properties.sheetId === sheetId).pop();
      return sheet.properties.gridProperties.rowCount;

      // const request = this.endpoints.spreadsheets.getByDataFilter();
      // request.addColumn(sheetId, 0, 1);
      // request.addQuery({fields: 'sheets.data.rowData.values.userEnteredValue,sheets.properties.sheetId'});
      // const json = request.fetch().json;
      // return json.sheets.filter(
      //   sh => sh.properties.sheetId === sheetId
      // ).pop().data.filter(
      //   d => d.userEnteredValue!==undefined
      // ).length;
    }

    getDimensions(sheetId) {
      return {
        column: this.getLastColumn(sheetId),
        row: this.getLastRow(sheetId)
      }
    }

    headerRequest(sheetId, header) {
      // use mixin pattern
      const searchRequest = this.endpoints.developerMetadata.search({header});
      // sheet location where key = headers and value is the sheet ID
      searchRequest.byLocationType('column', {metadataKey: header, metadataValue: sheetId.toString()});
      return searchRequest;
    }

    idRequest(sheetId, id) {
      const idString = id.toString();
      const searchRequest = this.endpoints.developerMetadata.search({id});  // mixin can be id
      // sheet location where key = headers and value is the sheet ID
      searchRequest.byLocationType('row', {metadataKey: idString, metadataValue: sheetId.toString()});
      return searchRequest;
    }

    newColumnMd(sheetId, header, idx) {
      const addHeaderReq = this.endpoints.spreadsheets.batchUpdate();
      addHeaderReq.createMetaData({
        metadataKey: header,
        metadataValue: sheetId.toString()
      }, {
        location: {
          dimensionRange: {
            sheetId: sheetId,
            dimension: "COLUMNS",
            startIndex: idx,
            endIndex: idx + 1
          }
        },
        visibility: "DOCUMENT"
      });
      const response = addHeaderReq.fetch();
      if (!response.ok) throw new Error(response.json.error.message);
      return response;
    }

    newRowMd(sheetId, id, idx) {
      const addRowReq = this.endpoints.spreadsheets.batchUpdate();
      addRowReq.createMetaData({
        metadataKey: id.toString(),
        metadataValue: sheetId.toString()
      }, {
        location: {
          dimensionRange: {
            sheetId: sheetId,
            dimension: "ROWS",
            startIndex: idx,
            endIndex: idx + 1
          }
        },
        visibility: "DOCUMENT"
      });
      const response = addRowReq.fetch();
      if (!response.ok) throw new Error(response.json.error.message);
      return response;
    }
  }

  class SheetsMetadataDoc {

    /**
     * Needs to also set up the sheet Id internally, as determined by metadata
     * if metadata by sheetKey not present, create it with that name
     * Responsbile for setting up sheet
     */
    constructor (id, sheetKey) {
      this.endpoints = GSheetEndpoints.fromId(id);
      this.search = new Search(this.endpoints);
      this.sheetId = this.search.getSheetId(sheetKey);
      Logger.log(`sheetID: ${this.sheetId}`);
    }

    static fromId(id, sheetKey) {
      return new SheetsMetadataDoc(id, sheetKey);
    }

    init() {

    }

    query({id}) {

    }

    update({json}) {

    }

    insert({json}) {

    }



    /** 
     * Takes a list of jsons, converts them to rows to see the headers, ensures those headers are present in md
     *   via batch call
     * Then goes through each json, and if present builds update request, if not, builds insert request.
     * Finally, batch calls the requests
     * @returns {Object} - the replies
     */
    apply({jsons, fields='totalUpdatedCells'}={}) {
      const Updater = this.endpoints.values.batchUpdateByDataFilter();
      Updater.setFields(fields);

      const ids = jsons.map(j => j.id || null);
      const rows = dottie.jsonsToRows(jsons);
      const headers = rows[0];

      const {column: lastColumnIdx, row: lastRowIdx} = this.search.getDimensions(this.sheetId);

      const checkHeaderBatch = GSheetEndpoints.batch();
      for (const header of headers) {
        const request = this.search.headerRequest(this.sheetId, header);
        checkHeaderBatch.add({request});
      }
      const checkIdsBatch = GSheetEndpoints.batch();
      for (const id of ids) {
        const request = this.search.idRequest(this.sheetId, id);
        checkIdsBatch.add({request});
      }

      // go through headers, making them if not present, count how many column we need to add
      // build colum map while we're at it
      const grid = {
        headers: new Map(),
        ids: new Map()
      }

      const storeResponseToGrid = (typ, res) => {
        // return 1 if empty, 0 if not, and store in grid
        const j = res.json;
        if (isEmpty(j)) {
          return 1;
        } else {
          const k = {'headers': res.request.header, 'ids': res.request.id}[typ];
          const i = j.matchedDeveloperMetadata.pop().developerMetadata.location.dimensionRange.startIndex;
          const t = {'headers': 'Column', 'ids': 'Row'}[typ];
          const v = {
            [`start${t}Index`]: i,
            [`end${t}Index`]: i + 1
          };
          grid[typ].set(k, v);
        }
        return 0;
      };

      const storeNewToGrid = (typ, key, idx) => {
        const t = {'headers': 'Column', 'ids': 'Row'}[typ];
        const v = {
          [`start${t}Index`]: idx,
          [`end${t}Index`]: idx + 1
        }
        grid[typ].set(key, v)
      };

      let countAddCols = 0;
      for (const response of checkHeaderBatch) {
        countAddCols += storeResponseToGrid('headers', response);
      }
      let countAddRows = 1;   // assume header TODO: don't assume header
      for (const response of checkIdsBatch) {
        countAddRows += storeResponseToGrid('ids', response);
      }

      // increase number of columns/rows if necessary
      if (countAddCols > 0) {
        this.search.setColumns(this.sheetId, lastColumnIdx + countAddCols - 1);
      }
      
      if (countAddRows > 1) {
        this.search.setRows(this.sheetId, lastRowIdx + countAddRows - 1);
      }

      let newHeaderIdx = lastColumnIdx - 1;
      // then go through them again, and add them with md request
      for (const response of checkHeaderBatch) {
        const json = response.json;
        if (isEmpty(json)) {
          // no location info for this header info yet, let's make it!
          const header = response.request.header
          this.search.newColumnMd(this.sheetId, header, newHeaderIdx);
          storeNewToGrid('headers', header, newHeaderIdx);

          // add it to the list of things to update, upon creation
          const gridRange = grid.headers.get(header);
          gridRange.startRowIndex = 0;
          gridRange.endRowIndex = 1;
          gridRange.sheetId = this.sheetId;
          Updater.addGridRange(gridRange, {values: [[header]]});
          
          newHeaderIdx += 1;
        } 
        // okay, save it
      }

      let newRowIdx = lastRowIdx - 1;
      // then go through them again, and add them with md request
      for (const response of checkIdsBatch) {
        const json = response.json;
        if (isEmpty(json)) {
          // no location info for this header info yet, let's make it!
          this.search.newRowMd(this.sheetId, response.request.id, newRowIdx);
          storeNewToGrid('ids', response.request.id, newRowIdx)
          newRowIdx += 1;
        } 
        // okay, save it
      }

      // keep a map of grid data for ids and columns
      // build gridRange with start/end row/column and sheetId
      // batchUpdateByDataFilter with valueInput = 'raw'
      for (const [ridx, row] of rows.slice(1).entries()) {
        for (const [hidx, value] of row.entries()) {
          const gridRange = {
            ...grid.headers.get(headers[hidx]),  // cols
            ...grid.ids.get(jsons[ridx].id),     // rows
            sheetId: this.sheetId         // sheetId
          };
          const values = [[rows[ridx+1][hidx]]];
          Updater.addGridRange(gridRange, {values});
        }
      }
      const replies = Updater.fetch().json;
      return replies;
    }

  }

  exports.SheetsMetadataDoc = SheetsMetadataDoc;

})(Import);
