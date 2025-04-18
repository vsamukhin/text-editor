import { EditorContent, useEditor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import { saveAs } from 'file-saver';
import * as mammoth from 'mammoth';
import { useEffect, useState } from 'react';
import * as XLSX from 'xlsx';
import { saveDocxFile } from '../utils/saveDocxFile';
import { saveTextFile } from '../utils/saveTextFile';


const TextEditor: React.FC = () => {
  const [content, setContent] = useState<string>('');

  const editor = useEditor({
    extensions: [StarterKit],
    content,
    onUpdate: ({ editor }) => {
      setContent(editor.getHTML());
    },
  });

  useEffect(() => {
    if (editor && content) {
      editor.commands.setContent(content);
    }
  }, [content, editor]);

  const handleTextUpload = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const text = e.target?.result as string;
      // Улучшенная обработка переносов строк
      const html = text
        .split(/\r?\n/)
        .map((line) => line.trim() ? `<p>${line}</p>` : '<p><br></p>')
        .join('');
      
      setContent(html);
    };
    reader.readAsText(file);
  };

  const handleExcelUpload = (file: File, editor: any) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });

      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const range = XLSX.utils.decode_range(sheet['!ref']!);
      const merges = sheet['!merges'] || [];

      const htmlRows: string[] = [];

      for (let row = range.s.r; row <= range.e.r; row++) {
        let rowHtml = '<tr>';

        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });

          // если это объединённая ячейка, не рендерим её снова
          if (merges.some(m => m.s.r < row && m.e.r >= row && m.s.c <= col && m.e.c >= col)) {
            continue;
          }

          const cell = sheet[cellAddress];
          const value = cell?.v ?? '';

          const merge = merges.find(m => m.s.r === row && m.s.c === col);
          const rowspan = merge ? merge.e.r - merge.s.r + 1 : 1;
          const colspan = merge ? merge.e.c - merge.s.c + 1 : 1;

          rowHtml += `<td${rowspan > 1 ? ` rowspan="${rowspan}"` : ''}${colspan > 1 ? ` colspan="${colspan}"` : ''}>${value}</td>`;
        }

        rowHtml += '</tr>';
        htmlRows.push(rowHtml);
      }

      const htmlTable = `<table border="1" style="border-collapse: collapse; width: 100%; font-size: 14px">${htmlRows.join('')}</table>`;

      // Вставка в Tiptap редактор
      editor.commands.insertContent(htmlTable);
    };

    reader.readAsArrayBuffer(file);
  };

  const handleDocxUpload = async (file: File) => {
    const reader = new FileReader();

    reader.onload = async (e) => {
      try {
        const arrayBuffer = e.target?.result as ArrayBuffer;

        const convertImage = (mammoth.images as any).inline((image: any) => {
          return image.read("base64").then((imageBuffer: string) => {
            return {
              src: `data:${image.contentType};base64,${imageBuffer}`,
              alt: image.altText || "",
            };
          });
        });

        const result = await mammoth.convertToHtml(
          { arrayBuffer },
          {
            styleMap: [
              "p[style-name='Heading 1'] => h1:fresh",
              "p[style-name='Heading 2'] => h2:fresh",
              "r[style-name='Strong'] => strong",
              "r[style-name='Emphasis'] => em",
              "p[style-name='List Paragraph'] => ul > li:fresh",
              "r[style-name='Highlight'] => mark"
            ],
            convertImage,
          }
        );

        let html = result.value;

        // Дополнительная очистка
        html = html
          .replace(/<p><\/p>/g, '<p><br></p>')
          .replace(/<p>\s*<\/p>/g, '<p><br></p>')
          .replace(/<strong><\/strong>/g, '')
          .replace(/<em><\/em>/g, '');

        setContent(html);
      } catch (error) {
        console.error("Error converting DOCX:", error);
        alert("Ошибка при загрузке DOCX файла");
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const handleSaveTxt = () => {
    if (!editor) return;
    
    // Получаем текст с сохранением форматирования
    const text = editor.getText({ blockSeparator: '\n\n' });
    saveTextFile(text, 'document.txt');
  };

  const handleSaveDocx = () => {
    if (!editor) return;
    
    // Получаем HTML содержимое
    const html = editor.getHTML();
    saveDocxFile(html, 'document.docx');
  };

  const handleSaveExcel = () => {
    if (!editor) return;

    const tempEl = document.createElement('div');
    tempEl.innerHTML = editor.getHTML();

    const tables = Array.from(tempEl.getElementsByTagName('table'));
    if (tables.length === 0) {
      alert('Нет таблиц для сохранения.');
      return;
    }

    const firstTable = tables[0];
    const rows = Array.from(firstTable.rows).map((row) =>
      Array.from(row.cells).map((cell) => cell.innerText)
    );

    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    const wbout = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    saveAs(blob, 'table.xlsx');
  };

  return (
    <div>
      <input
        type="file"
        accept=".txt,.docx, .xlsx"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) {
            if (file.name.endsWith('.txt')) {
              handleTextUpload(file);
            } else if (file.name.endsWith('.docx')) {
              handleDocxUpload(file);
            } else if (file.name.endsWith('.xlsx')) {
              handleExcelUpload(file, editor);
            }
            e.target.value = '';
          }
        }}
      />
      <div style={{ border: '1px solid #ccc', padding: 10, marginTop: 10 }}>
        <div className="toolbar" style={{ marginBottom: 10 }}>
          <button 
            onClick={() => editor?.chain().focus().toggleBold().run()}
            style={{ fontWeight: editor?.isActive('bold') ? 'bold' : 'normal' }}
          >
            Bold
          </button>
          <button 
            onClick={() => editor?.chain().focus().toggleItalic().run()}
            style={{ fontStyle: editor?.isActive('italic') ? 'italic' : 'normal' }}
          >
            Italic
          </button>
          <button 
            onClick={() => editor?.chain().focus().toggleBulletList().run()}
            style={{ fontWeight: editor?.isActive('bulletList') ? 'bold' : 'normal' }}
          >
            List
          </button>
          <button 
            onClick={() => editor?.chain().focus().toggleHeading({ level: 1 }).run()}
            style={{ fontWeight: editor?.isActive('heading', { level: 1 }) ? 'bold' : 'normal' }}
          >
            H1
          </button>
          <button 
            onClick={() => editor?.chain().focus().toggleHeading({ level: 2 }).run()}
            style={{ fontWeight: editor?.isActive('heading', { level: 2 }) ? 'bold' : 'normal' }}
          >
            H2
          </button>
        </div>   
        <div style={{ width: '100%', height: '70vh', overflow: 'hidden' }}>
          <EditorContent editor={editor}/>
        </div>
        <div style={{ marginTop: 20 }}>
          <button onClick={handleSaveTxt}>💾 Сохранить .txt</button>
          <button onClick={handleSaveDocx}>💾 Сохранить .docx</button>
          <button onClick={handleSaveExcel}>💾 Сохранить .xlsx</button>
        </div>
      </div>
    </div>
  );
};

export default TextEditor;