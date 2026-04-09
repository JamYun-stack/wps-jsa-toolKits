# fileSystemUtils.js Reference

Use this reference when editing `../../fileSystemUtils.js` and `../../fileSystemUtils.md`.

## Current Public Surface

- Picker helpers: `openFolderPicker`, `openFilePicker`
- Existence checks: `fileExists`, `folderExists`
- Path helpers: `normalizePath`, `joinPath`, `getPathName`, `getParentFolderPath`, `getFileBaseName`, `getFileExtend`, `changeFileExtend`
- Folder creation: `createFolder`, `ensureFolder`, `ensureParentFolder`
- File metadata: `getTempFolderPath`, `getFileSize`, `getFileModifiedTime`
- File operations: `copyFile`, `moveFile`, `deleteFile`
- Folder operations: `copyFolder`, `moveFolder`, `deleteFolder`
- Directory listing: `getFilesByPath`, `listFilesByPath`, `getFoldersByPath`, `listFoldersByPath`

## Implementation Expectations

- Prefer WPS/JSA-compatible functions before Windows objects.
- Keep `ActiveXObject` usage as fallback unless no WPS/JSA equivalent exists, such as recursive folder copy.
- Keep root-path protections on destructive folder operations.
- Preserve the return format of `getFilesByPath`:

```js
{
    "demo.xlsx": {
        fileName: "demo.xlsx",
        path: "D:\\test\\demo.xlsx",
        extend: "xlsx"
    }
}
```

## Validation Checklist

- Run `node --check .\fileSystemUtils.js`.
- Search for stale names such as `publicModule02` and `_pm02`.
- Confirm every function has Chinese JSDoc with `@param` entries for named parameters and an `@returns` entry.
- Update `fileSystemUtils.md` in the same change when public behavior changes.
