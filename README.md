# planning-vba-automation

## Configuration

The VBA modules expect a worksheet named `Config` containing the base path for PDF exports.
Specify your OneDrive folder in cell **B2** of this sheet (e.g. `C:\Users\myname\OneDrive\`).
If the cell is empty, the code falls back to trying a few hard coded user names.
