# GSheetsMetadataDoc

Keep a google sheet synced with a data source represented by a list of jsons (which all have the property `id`), such as provided by API endpoints. Users can rearrange columns, change name of header, columns or rows — format it in any way they choose – and subsequent updates will stay in sync.

## Getting Started

Add the library `1kMavLN2B4SyluEQUjbWCDx5-vBkABKbgWkKtBX5PnaSJ9JZKTs4g3MVw`, and you initialize like this:

```js
const doc = GSheetsMetadataDoc.fromId(ssId, someKey);
```

The `ssId` is the ID of the spreadsheet, and `someKey` is internally used to ensure sync occurs correctly. Both these values need to remain the same for subsequent updates to work.

You'll need a list of jsons to pass to the `update` method:

```js
const jsons = [{id: 1, values: [0, 1]}, ...];
doc.update({jsons});
```

## Motivation

Spreadsheets are really useful, and they could be more useful if there was a way to keep them updated as above.
