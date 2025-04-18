import { saveAs } from 'file-saver';

export const saveTextFile = (htmlContent: string, fileName: string) => {
  const plainText = htmlContent.replace(/<[^>]+>/g, '');
  const blob = new Blob([plainText], { type: 'text/plain;charset=utf-8' });
  saveAs(blob, fileName.endsWith('.txt') ? fileName : `${fileName}.txt`);
};
