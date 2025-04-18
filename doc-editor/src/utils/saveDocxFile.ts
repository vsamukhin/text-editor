import { saveAs } from 'file-saver';
import htmlDocx from 'html-docx-js/dist/html-docx';

export const saveDocxFile = (htmlContent: string, fileName: string) => {
  const html = `<!DOCTYPE html>
  <html>
    <head>
      <meta charset="utf-8">
    </head>
    <body>${htmlContent}</body>
  </html>`;

  const blob = htmlDocx.asBlob(html);
  saveAs(blob, fileName.endsWith('.docx') ? fileName : `${fileName}.docx`);
};
