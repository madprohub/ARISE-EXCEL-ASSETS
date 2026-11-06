# ARISE v9 Premium Animation Assets

Official set selected: **ARISE_v9_PREMIUM_NEON_ASSETS**, 18 JPEG frames (`frame_00.jpg` to `frame_17.jpg`).

**Status:** Binary frames are pending upload. The source files have been prepared, but this GitHub connector does not accept local binary files directly.

## Upload
Extract the provided ARISE_PREMIUM_GITHUB_UPLOAD_PACKAGE.zip and upload the JPG files to `assets/premium/` on this repository's main branch. Verify that [frame_00.jpg](https://raw.githubusercontent.com/madprohub/ARISE-EXCEL-ASSETS/main/assets/premium/frame_00.jpg) works before running the loader.

## Windows VBA cache loader
See `vba/ARISE_GitHub_AssetCache.bas` and `docs/PREMIUM_INSTALL.txt`. The cache is under `%LOCALAPPDATA%\MADPRO\ARISE\PREMIUM_NEON\`; it is **not** embedded in the workbook. Requires initial internet access and macros enabled. Do not delete the original frame folder before testing.
