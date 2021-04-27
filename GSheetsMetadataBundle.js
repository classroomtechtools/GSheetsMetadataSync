const Import = Object.create(null);
(function (exports) {

  const isEmpty = (obj) => Object.keys(obj).length === 0;
  const check_error = (resp, comment="") => {
    if (!resp.ok) throw new Error(`${resp.json.error.message} ${comment}`);
    return resp;
  };

  class Search {
    /**
     * returns sheet id by sheetKey
     * if not present, return null
     */
    constructor (endpoints, visibility='PROJECT') {
      this.endpoints = endpoints;
      this.countRows = 0;
      this.countColumns = 0;
      this.visibility = visibility;
    }

    /**
     * Returns the sheetId stored in md, or create it if not present
     * TODO: Add check if tab already exists when creating
     */
    getSheetId(sheetKey) {
      let sheetId, sheetName;

      const metadataKey = `ctt_sheetKey_${sheetKey}`;
      const request = this.endpoints.developerMetadata.search();
      request.bySpreadsheetLoc({metadataKey, visibility: this.visibility});
      const json = check_error(
        request.fetch(), 'looking for sheetKey'
      ).json;

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
        const addSheetResp = check_error(
                                addSheetReq.fetch(), 'when trying to create sheet'
                             );
        const reply = addSheetResp.json.replies.pop();
        // got it:
        sheetId = reply.addSheet.properties.sheetId;
        sheetName = reply.addSheet.properties.title;

        // now create sheet md
        const addSheetMdReq = this.endpoints.spreadsheets.batchUpdate();
        addSheetMdReq.createMetaData({
          metadataKey, 
          metadataValue: sheetId.toString(),
        }, {
          location: {
            spreadsheet: true
          },
          visibility: this.visibility
        });
        const addSheetMdJson = check_error(
          addSheetMdReq.fetch(), 'when trying to create new sheet metadata'
        ).json;
      } else {
        sheetId = parseInt(json.matchedDeveloperMetadata.pop().developerMetadata.metadataValue);

        const ssreq = this.endpoints.spreadsheets.get();
        const ssresp = check_error(
          ssreq.fetch(), 'while getting spreadsheet'
        );
        const ssjson = ssresp.json;

        const filteredSheet = ssjson.sheets.filter(
          sh => sh.properties.sheetId === sheetId
        );
        if (filteredSheet.length === 0) {
          throw new Error(`The old sheet with sync key ${sheetKey} must have been deleted. Try a new key?`);
        }
        sheetName = filteredSheet.pop().properties.title;
      }
      return {sheetId, sheetName};
    }

    getSpreadsheetProperties() { 
      return this.endpoints.spreadsheets.get().fetch().json.properties;
    }

    getDimensions(sheetId) {
      const request = this.endpoints.spreadsheets.get();
      const json = request.fetch().json;
      const sheet = json.sheets.filter(sh => sh.properties.sheetId === sheetId).pop();
      const locale = json.locale;
      const timeZone = json.timeZone;
      const frozenRowCount = sheet.properties.gridProperties.frozenRowCount;

      // One off error potential here:
      const rowCount = sheet.properties.gridProperties.rowCount;
      const columnCount = sheet.properties.gridProperties.columnCount;

      const incRowCount = () => {
        this.countRows += 1;
      };

      const incColumnCount = () => {
        this.countColumns += 1;
      };

      const updateRowColumnCount = () => {
        const UpdateRowsColumnsRequest = this.endpoints.spreadsheets.batchUpdate(/*  */);
        const {x, y} = {
          x: {l: 'row', v: this.countRows + rowCount, i: rowCount},
          y: {l: 'column', v: this.countColumns + columnCount, i: columnCount}
        };
        for (const {l, v, i} of [x, y]) {
          if (i > 0) {
            UpdateRowsColumnsRequest.updateSheetProperties(sheetId, {
              gridProperties: {
                [`${l}Count`]: v
              }
            }, `gridProperties.${l}Count`);
          }
        }

        if ( UpdateRowsColumnsRequest.payload.requests.length > 0 ) {
          const resp = check_error(
            UpdateRowsColumnsRequest.fetch(), 'while updating rows, columns'
          );
          
        }
      };

      return {
        rowCount, columnCount, locale, timeZone,
        incRowCount, incColumnCount,
        updateRowColumnCount,
        frozenRowCount
      };
    }

    headerRequest(batch, sheetId, header) {
      // use mixin pattern yikes!
      //const searchRequest = this.endpoints.developerMetadata.search({header});
      // sheet location where key = headers and value is the sheet ID
      batch.byLocationType('column', sheetId, {
        metadataKey: header,
        visibility: this.visibility
      });
    }

    idRequest(batch, sheetId, id) {
      const idString = id.toString();
      //const searchRequest = this.endpoints.developerMetadata.search({id});  // mixin can be id
      // sheet location where key = headers and value is the sheet ID
      batch.byLocationType('row', sheetId, {metadataKey: idString, visibility: this.visibility});
    }

    newColumnMd(batch, sheetId, header, idx) {
      batch.createMetaData({
        metadataKey: header
      }, {
        location: {
          dimensionRange: {
            sheetId: sheetId,
            dimension: "COLUMNS",
            startIndex: idx,
            endIndex: idx + 1
          }
        },
        visibility: this.visibility
      });
    }

    newRowMd(batch, sheetId, id, idx) {
      batch.createMetaData({
        metadataKey: id.toString()
      }, {
        location: {
          dimensionRange: {
            sheetId: sheetId,
            dimension: "ROWS",
            startIndex: idx,
            endIndex: idx + 1
          }
        },
        visibility: this.visibility
      });
    }
  }

  class SheetsMetadataDoc {

    /**
     * Needs to also set up the sheet Id internally, as determined by metadata
     * if metadata by sheetKey not present, create it with that name
     * Responsbile for setting up sheet
     */
    constructor (id, sheetKey, visibility='PROJECT') {
      this.id = id;
      this.endpoints = GSheetEndpoints.fromId(id);
      this.visibility = visibility;
      this.search = new Search(this.endpoints, this.visibility);
      const {sheetId, sheetName} = this.search.getSheetId(sheetKey);
      this.sheetId = sheetId;
      this.sheetName = sheetName;
    }

    static fromId(id, sheetKey, {visibility="PROJECT"}={}) {
      return new SheetsMetadataDoc(id, sheetKey, visibility);
    }

    getTimeZone () {
      return this.search.getSpreadsheetProperties().timeZone;
    }

    getSpreadsheetMetadata({metadataKey=null, metadataValue=null}) {
      const request = this.endpoints.developerMetadata.search();
      request.bySpreadsheetLoc({metadataValue, metadataKey, visibility: this.visibility});
      return request.fetch().json.matchedDeveloperMetadata;
    }

    createSpreadsheetMetadata({metadataKey=null, metadataValue=null}) {
      const request = this.endpoints.spreadsheets.batchUpdate();
      const md = {};
      if (metadataKey) md.metadataKey = metadataKey;
      if (metadataValue) md.metadataValue = metadataValue;
      request.createMetaData(md, {
        location: {
          spreadsheet: true
        },
        visibility: this.visibility
      });
      return check_error(
        request.fetch(), 'when creating spreadsheet metadata with metadataKey = ' + metadataKey + ' and metadataValue = ' + metadataValue
      ).json;
    }

    updateSpreadsheetMetadata(metadataId, {metadataValue=null, metadataKey=null, visibility=null, location=null}) {
      const request = this.endpoints.spreadsheets.batchUpdate();
      const md = {};
      if (metadataKey) md.metadataKey = metadataKey;
      if (metadataValue) md.metadataValue = metadataValue;
      request.updateMetaData(metadataId, md, {visibility, location});
      return check_error(
        request.fetch(), 'when updating spreadsheet metadata with metadataKey = ' + metadataKey + ' and metadataValue = ' + metadataValue
      ).json;
    }

    /** 
     * @params {Object} jsons - A list of json
     * @params {String} fields - what to include
     * @params {String[]} priorityHeaders - flush left
     * @params {Boolean} isIterative - indicates this is not a wholesale update, so handle jsons differently
     * @returns {Object} - the replies
     */
    apply({jsons, fields='totalUpdatedCells', priorityHeaders=['id'], isIncremental=false,
            sortCallback = (a, b) => a.id - b.id})
    {
      if (jsons.length === 0) {
        return {totalUpdatedCells: 0, message: 'zero length json array'};
      }

      if ( jsons.filter(json => json.id == null).length > 0 ) {
        throw new Error(
          "All jsons must have id property which must remain stable (i.e. from a serialized database"
        );
      }

      // sort them by id (or however it is passed)
      jsons.sort( sortCallback );

      const {DateFns} = Import;

      const ValuesUpdater = this.endpoints.values.batchUpdateByDataFilter({
        valueInputOption: "USER_ENTERED"
      });
      ValuesUpdater.setFields(fields);

      const ShiftUpdater = this.endpoints.spreadsheets.batchUpdate();

      const ids = jsons.map(j => j.id || null);

      let characters = 0;

      const grid = {
        headers: new Map(),
        ids: new Map()
      };

      const storeMdToGrid = (md) => {
        const type = md.location.locationType;
        const typ = { 'COLUMN': 'headers', "ROW": 'ids'}[type];

        const k = md.metadataKey;  // note: this will alwyas be a string
        characters += k.length;
        const i = md.location.dimensionRange.startIndex;
        const t = {'headers': 'Column', 'ids': 'Row'}[typ];
        const v = {
          [`start${t}Index`]: i,
          [`end${t}Index`]: i + 1
        };
        grid[typ].set(k, v);
      };

      const storeNewToGrid = (md, idx) => {
        const type = md.locationType;
        const typ = {"COLUMN": "headers", "ROW": "ids"}[type];
        
        const k = md.metadataKey;
        characters += k.length;
        
        // i
        const t = {'headers': 'Column', 'ids': 'Row'}[typ];
        const v = {
          [`start${t}Index`]: idx,
          [`end${t}Index`]: idx + 1
        }
        grid[typ].set(k, v);
      };

      // don't delete "empty" or null values because we may need to
      let options = [];
      if (isIncremental) {
        options = [false, false];
      }

      const rows = dottie.jsonsToRows(jsons, priorityHeaders, ...options);
      const headers = rows[0];

      const Dimensions = this.search.getDimensions(this.sheetId);

      // build the header requests we need
      const CheckHeaderBatch = this.endpoints.developerMetadata.search();
      for (const header of headers) {
        this.search.headerRequest(CheckHeaderBatch, this.sheetId, header);
      }
      const CheckIdsBatch = this.endpoints.developerMetadata.search();
      for (const id of ids) {
        this.search.idRequest(CheckIdsBatch, this.sheetId, id);
      }

      // go through headers, making them if not present, count how many column we need to add
      // build colum map while we're at it
      

      const checkHeaderBatchResponsesJson = check_error(
        CheckHeaderBatch.fetch(), 'while getting headers'
      ).json;
      const checkIdsBatchResponsesJson = check_error(
        CheckIdsBatch.fetch(), 'while getting ids'
      ).json;

      const headerMetadatas = (checkHeaderBatchResponsesJson.matchedDeveloperMetadata || [])
                                        .map(m => m.developerMetadata);
      const idMetadatas = (checkIdsBatchResponsesJson.matchedDeveloperMetadata || [])
                                      .map(m => m.developerMetadata);
      // store grid info
      for (const metadata of [...headerMetadatas, ...idMetadatas]) {
        storeMdToGrid(metadata);
      }

      // now let's determine which headers and ids are missing
      // and add them with md request
      const ColumnRowUpdater = this.endpoints.spreadsheets.batchUpdate();
      //    ^--- this will need to be fetched first before ValuesUpdater updater
      //    for metadata info to be valid

      let allColumnIndexes;
      if (isEmpty(checkHeaderBatchResponsesJson)) {
        allColumnIndexes = [-1]; //.map( md => md.developerMetadata.location.dimensionGrid.startIndex);
      } else {
        allColumnIndexes = checkHeaderBatchResponsesJson.matchedDeveloperMetadata.map(
          md => md.developerMetadata.location.dimensionRange.startIndex
        );
      }
      let newColumnIndex = Math.max(...allColumnIndexes) + 1;

      // we have to go through each one we found in the data
      // but only necessary if we are going to add headers
      // isIncremental true means we skip this process
      if (!isIncremental) {
        for (const data of CheckHeaderBatch.payload.dataFilters) {
          const metadata = data.developerMetadataLookup;
          const header = metadata.metadataKey;

          if ( !grid.headers.has(header) ) {
            // no metadata info for this header info yet, let's make it!

            if (newColumnIndex >= Dimensions.columnCount) {
              Dimensions.incColumnCount(/* for subsequent Dimenisons.update call */);
            }
            this.search.newColumnMd(ColumnRowUpdater, this.sheetId, header, newColumnIndex);
            storeNewToGrid(metadata, newColumnIndex);   // pass literal ones instead

            // add it to the list of things to update, upon creation
            // but only if we have frozen rows on there
            const gridRange = grid.headers.get(header);
            if (Dimensions.frozenRowCount > 0) {
              gridRange.startRowIndex = 0;
              gridRange.endRowIndex = 1;
              gridRange.sheetId = this.sheetId;
              ValuesUpdater.addGridRange(gridRange, {values: [[header]]});
            }

            // put a request in to move it if it's a priority header
            if (priorityHeaders.includes(header)) {
              ShiftUpdater.moveDimension(
                this.sheetId, gridRange.startColumnIndex, priorityHeaders.indexOf(header)
              );
            }

            newColumnIndex += 1;
          } 
        }
      }  // end if !isIterative


      // determine which IDs are missing
      let allRowIndexes;
      if (isEmpty(checkIdsBatchResponsesJson)) {
        allRowIndexes = [0];
      } else {
        allRowIndexes = checkIdsBatchResponsesJson.matchedDeveloperMetadata.map(
          md => md.developerMetadata.location.dimensionRange.startIndex
        );
      }
      let newRowIndex = Math.max(...allRowIndexes) + 1;
      for (const data of CheckIdsBatch.payload.dataFilters) {
        const metadata = data.developerMetadataLookup;
        const id = metadata.metadataKey;

        if (!grid.ids.has(id)) {
          this.search.newRowMd(ColumnRowUpdater, this.sheetId, id, newRowIndex);
          storeNewToGrid(metadata, newRowIndex);
          if (newRowIndex >= Dimensions.rowCount) {
            Dimensions.incRowCount(/* for subsequent Dimenisons.update call */);
          }

          newRowIndex += 1;
        }
      }

      // add those new headers and row metadata now...
      // but first need to make sure we have enough rows and columns

      
      Dimensions.updateRowColumnCount(/* ensure there is room */);

      // now actually add the row and column md
      if (ColumnRowUpdater.payload.requests.length > 0) {
        const updated = check_error(
          ColumnRowUpdater.fetch(), 'while adding new header and row metadata'
        );
      }

      const re = new RegExp(/\d{4}-[01]\d-[0-3]\dT[0-2]\d:[0-5]\d:[0-5]\d\.\d+([+-][0-2]\d:[0-5]\d|Z)/);

      /**
       * converts a json date into a user-entered version that the spreadsheet can format to date
       * in the spreadsheet's local timezone!
       */
      const mapValue = (v) => {
        if (v == null) return '';   // ensure that any nulls are reverted to blank cells
        if (typeof v === 'object' && Object.keys(v).length === 0) return '';
        if (['number', 'boolean'].includes(typeof v)) return v;
        if (Array.isArray(v) && v.length === 0) return '';  // ensure to blank it for empty arrays
        if ( typeof v === 'string' && ['+'].includes(v[0]) ) return `'${v}`;  // phone numbers
        if (typeof v === 'string' && !re.test(v)) return v;
        const {tz} = Dimensions;
        let date;
        if (typeof v === 'string') {
          try {
            date = DateFns.parseISO(v);
          } catch (e) {
            date = v;
          }
        } else {
          date = v;
        }
        const zonedTime = DateFns.utcToZonedTime(date, tz);
        const formatted = DateFns.format(zonedTime, 'yyyy-MM-dd HH:mm:ss', {tz});
        return formatted;
      }

      // keep a map of grid data for ids and columns
      // build gridRange with start/end row/column and sheetId
      // batchUpdateByDataFilter with valueInput = 'raw'
      for (const [ridx, row] of rows.slice(1).entries()) {
        for (const [hidx, value] of row.entries()) {
          const gridRange = {
            ...grid.headers.get(headers[hidx]),  // cols
            ...grid.ids.get(jsons[ridx].id.toString()),     // rows
            sheetId: this.sheetId         // sheetId
          };
          //const values = [[rows[ridx+1][hidx]]];
          const values = [[mapValue(value)]];
          ValuesUpdater.addGridRange(gridRange, {values});
        }
      }

      // ready to write
      const reply = check_error(
        ValuesUpdater.fetch(), 'while updating values'
      ).json;

      if (ShiftUpdater.payload.requests.length > 0) {
        SpreadsheetApp.flush();  // wait to make sure updates occur
        const moved = check_error(
          ShiftUpdater.fetch(), 'while moving columns'
        );
      }
      
      return reply;
    }

  }

  exports.SheetsMetadataDoc = SheetsMetadataDoc;

})(Import);
