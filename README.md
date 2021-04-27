# GSheetsMetadataSync

Keep a google sheet synced with a data source represented by a list of jsons (which all have the property `id`), such as provided by API endpoints. Users can rearrange columns, change name of header, columns or rows — format it in any way they choose – and subsequent updates will stay in sync.

It uses metadata to track the rows (jsons with `id`) and the columns (nested properties in the json).

```
const jsons = [
  {
    id: 1,
    values: [0, 1],
    nested: {
      value: 'hello',
      obj: {
        arr: [
          'one', 'two'
        ]
      }
    }
  }, 
  ...
];
```

will result in a tab with the following columns

| id | values[0] | values[1] | nested.obj.arr[0] | nested.obj.arr[1] | nested.value |
| ------- | ----------- | --- | ----------- | --- | ----------- |
| 1 | 0 | 1 | one | two | hello |
...

## Getting Started

Add the library `1kMavLN2B4SyluEQUjbWCDx5-vBkABKbgWkKtBX5PnaSJ9JZKTs4g3MVw`, and you initialize like this:

```js
const doc = GSheetsMetadataDoc.fromId(ssId, someKey);
```

The `ssId` is the ID of the spreadsheet, and `someKey` is internally used to ensure sync occurs correctly. Both these values need to remain the same for subsequent updates to work.

You'll need a list of jsons to pass to the `apply` method:

```js
const jsons = [{id: 1, values: [0, 1]}, ...];
doc.apply({jsons});
```

It will then create a new tab (with default name the value of `someKey`) and populate the sheet in rows for each json, with columns across for each field encountered.

If you pass `isIncremental` as `true`, you're letting the sync understand that you're just doing simple updates, and do not create new columns encountered.

```js
doc.apply({jsons, isIncremental: true});
```

Function is defined as thus:

```js
 /** 
   * @params {Object[]} jsons - A list of json
   * @params {String} fields - what fields to include in the batchUpdateByDataFilter response 
   * @params {String[]} priorityHeaders - headers to flush left
   * @params {Boolean} isIterative - indicates this is not a wholesale update, so handle jsons differently
   * @returns {Object} - the replies
   */
  apply({jsons, fields='totalUpdatedCells', priorityHeaders=['id'], isIncremental=false,
          sortCallback = (a, b) => a.id - b.id})
{ }
```

## Motivation

Spreadsheets are really useful, and they could be more useful if there was a way to keep them updated as above.
