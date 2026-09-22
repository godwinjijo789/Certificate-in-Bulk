# Certify — Certificate Generator

A React + Vite website that imports participant names from Excel, lets you position and style the name over a certificate image, and exports individual PNGs in a ZIP.

## Requirements
- Node.js 20 or newer (LTS recommended)
- VS Code

## Run locally
1. Extract the ZIP.
2. Open the extracted `certificate-generator` folder in VS Code.
3. Open Terminal → New Terminal.
4. Run `npm install`
5. Run `npm run dev`
6. Open the local URL shown by Vite (usually `http://localhost:5173`).

## Excel format
Use the first worksheet. Row 1 must contain column headings, such as `Name`, `Email`, `Department`. Each participant goes on a separate row. The app tries to select a column containing “name”; you can change it after import.

## Use
1. Upload a PNG/JPG/WebP certificate template (best results: landscape, high-resolution image).
2. Upload an `.xlsx` workbook.
3. Choose the name column.
4. Click the certificate preview to set the name position; fine-tune with sliders.
5. Adjust font, size, color, and bold.
6. Click **Generate & download ZIP**.

## Notes
- This version accepts image templates (PNG/JPG/WebP), not PDF templates.
- Output certificates are PNG images, bundled in a ZIP.
- Processing happens in your browser; the app has no backend or login/database.
- For `.xls`, save the workbook as `.xlsx` first.
