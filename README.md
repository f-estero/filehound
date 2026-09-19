# FileHound

FileHound is a Windows desktop utility built for one practical goal: finding files reliably when Windows Search is slow, incomplete, or not working well.

It is designed around a strict safety rule:

- source files are never modified;
- copy operations only write to user-selected output folders;
- name collisions in the destination are automatically renamed, never overwritten.

Repository: [f-estero/filehound](https://github.com/f-estero/filehound.git)

## Features

### Single search

Search through a folder and all subfolders using:

- `Exact`
- `Contains`
- `Similar`

From the results view you can:

- open a file;
- open the containing folder;
- send results to the Safe Copy tab.

### Batch search from Excel or CSV

Load:

- `xlsx`
- `xls`
- `csv`

Choose the column that contains the requested file names and run a multi-file search in one pass.

### Safe Copy

The Safe Copy tab lets you:

- copy only selected results;
- copy all found results;
- preserve original files untouched;
- avoid overwriting destination files by automatically generating unique names.

### Roe Photos / Metadata

The Roe-specific area is kept separate from the general finder workflow and includes:

- Roe-style variant matching such as `name.jpg`, `name__1.jpg`, `name__2.jpg`;
- Roe result export with GPS metadata columns;
- safe copy for Roe matches;
- GPS metadata check on output copies only.

## Safety model

FileHound is intentionally conservative:

- the source side is read-only from the app's point of view;
- all writes happen only in output folders;
- GPS metadata is written only to copied files, never to originals;
- duplicate names in the destination are renamed automatically.

## Requirements

- Windows
- Python 3.10+

Install dependencies:

```bash
pip install -r requirements.txt
```

## Run

```bash
python filehound.py
```

## Project structure

The current app is split into four functional areas:

- `Search`
- `Batch Excel / CSV`
- `Safe Copy`
- `Roe Photos / Metadata`

This structure reflects the product direction:

- the main value is a safe file finder for Windows;
- Roe and GPS operations are specialized modules, not the main workflow.

## Author

- Author: `f-estero`
- Project name: `FileHound`

## Future improvements

- optional local indexing for faster repeated searches;
- filters by extension, size, and modified date;
- checkbox selection directly in result tables;
- saved recent searches;
- persistent favorite folders;
- packaged Windows executable release.
