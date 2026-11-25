# Manga Compressor

A simple Windows PowerShell tool that combines multiple manga chapter files into CBZ archives for e-readers.

Choose between combining all chapters into one large CBZ archive or automatically splitting them into separate archives by volume.

## Features

- **Merge all chapters** into one CBZ archive.
- **Split by volumes** into separate CBZ archives.
- Supports CBZ and ZIP chapter archives.
- Supports folders containing manga page images.
- Natural sorting of chapters and pages.
- Built-in progress bar.
- Automatic output folder creation.
- Prevents overwriting existing archives.

## Requirements

- Windows 10 / Windows 11.
- Windows PowerShell 5.1 or PowerShell 7+.
- No Python or additional dependencies required.

## How to use

1. Download and extract the ZIP package.
2. Run the `.bat` launcher.
3. Select the folder containing your manga chapters.
4. Choose an export mode:
   - `1` - Merge everything into one CBZ archive.
   - `2` - Split chapters into separate CBZ archives by volume.
5. Enter the manga title.
6. Wait for the process to finish.

All generated archives will be saved in a `CBZ_Export` folder.

## Volume Detection

The script detects volume numbers from archive filenames and parent folder names.

Supported naming formats:

- `Tom 1`
- `Том 2`
- `Volume 3`
- `Vol. 4`
- `v05`

Chapters without a detected volume number are placed into volume 1.

**Note:** Volume detection relies on filenames and folder names. The script does not determine volume boundaries from chapter numbers or manga content.

## Supported Formats

Input archives:
- `.cbz`
- `.zip`

Image formats:
- `.jpg`
- `.jpeg`
- `.png`
- `.webp`
- `.gif`
- `.bmp`
- `.avif`
- `.tif`
- `.tiff`

## Output Example

```text
CBZ_Export/
├── Manga-tom-1.cbz
├── Manga-tom-2.cbz
└── Manga-tom-3.cbz
```

The resulting CBZ files can be transferred to any e-reader that supports the CBZ format.

## Disclaimer

This is an independent community project and is not affiliated with any manga reader or content provider.
