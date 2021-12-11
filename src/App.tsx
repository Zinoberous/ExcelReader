import React from 'react';
import './App.css';

import * as Excel from 'exceljs';

function toBuffer(ab:ArrayBuffer) {
  var buf = Buffer.alloc(ab.byteLength);
  var view = new Uint8Array(ab);
  for (var i = 0; i < buf.length; ++i) {
      buf[i] = view[i];
  }
  return buf;
}

function read(e:any) {
  const reader = new FileReader();
  reader.readAsArrayBuffer(e.target.files[0]);
  reader.onloadend = (ev) => {
    const arrBuffer:ArrayBuffer = ev.target?.result as ArrayBuffer;
    new Excel.Workbook().xlsx.load(toBuffer(arrBuffer)).then((workbook) => {
      // const worksheetsPool:string[] = ['Jan','Feb','Mär','Apr','Mai','Jun','Jul','Aug','Sep','Okt','Nov','Dez'];
      workbook.worksheets
      // .filter(worksheet => worksheetsPool.includes(worksheet.name))
      .forEach(worksheet => {
        const cellCount:number = worksheet.getRow(1).cellCount;

        const columns:string[] = [];
        for (let i = 1; i <= cellCount; i++) {
          const colValue:any = worksheet.getRow(1).getCell(i).value;
          
          let colName:string = '';
          if (colValue.richText) {
            colValue.richText.forEach((rtf:any,idx:number) => {
              if (idx !== 0) { colName += ' '; }
              colName += rtf.text.trim();
            });
          } else {
            colName = colValue;
          }

          columns[i] = colName;
        }

        const rows:Record<string,any>[] = [];
        worksheet.getRows(2,worksheet.rowCount)?.forEach((row,ri) => {
          if (row.cellCount > 0) {
            rows[ri] = {};
            for (let ci = 1; ci <= cellCount; ci++) {
              let cell = row.getCell(ci).master;

              const colKey:string = columns[ci];
              rows[ri][colKey] = cell.value;
              rows[ri][colKey + ' Note'] = (
                cell.note && (cell.note as any).texts && (cell.note as any).texts[0] && (cell.note as any).texts[0].text
                  ? (cell.note as any).texts[0].text
                  : ''
              );

              // // TODO: cell.comments doesn't exits, cell.note only inclide notes on cell.tnote.texts[0].text if using comments instead cell.note.texts is an empty array
              // rows[ri][colKey + ' Chat'] = cell.comments.map(comment => { return { user:comment.author, date:comment.created, message:comment.value }; });
            }
          }
        });

        console.log({
          worksheet:worksheet.name,
          columns,
          rows
        });
      });
    });
  }
}

function App() {
  return (
    <div className="App">
      <input id='worklog' type='file' onChange={read} />
    </div>
  );
}

export default App;
