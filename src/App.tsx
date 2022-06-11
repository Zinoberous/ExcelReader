import React from 'react';
import './App.css';

import * as Excel from 'exceljs';

function toBuffer(ab: ArrayBuffer) {
  var buf = Buffer.alloc(ab.byteLength);
  var view = new Uint8Array(ab);
  for (var i = 0; i < buf.length; ++i) {
    buf[i] = view[i];
  }
  return buf;
}

const workbook: any[] = [];

function read(e: any) {
  const reader = new FileReader();
  reader.readAsArrayBuffer(e.target.files[0]);
  reader.onloadend = (ev) => {
    const arrBuffer: ArrayBuffer = ev.target?.result as ArrayBuffer;
    new Excel.Workbook().xlsx.load(toBuffer(arrBuffer)).then((wb) => {
      // const worksheetsPool:string[] = ['Jan','Feb','Mär','Apr','Mai','Jun','Jul','Aug','Sep','Okt','Nov','Dez'];
      wb.worksheets
        // .filter(worksheet => worksheetsPool.includes(worksheet.name))
        .forEach(ws => {
          const cellCount: number = ws.getRow(1).cellCount;

          const columns: string[] = [];
          for (let i = 1; i <= cellCount; i++) {
            const colValue: any = ws.getRow(1).getCell(i).value;

            let colName: string = '';
            if (colValue && colValue.richText) {
              colValue.richText.forEach((rtf: any, idx: number) => {
                if (idx !== 0) { colName += ' '; }
                colName += rtf.text.trim();
              });
            } else {
              colName = colValue;
            }

            columns[i] = (colName || '').replace('ä', 'ae').replace('ö', 'oe').replace('ü', 'ue');
          }

          const rows: Record<string, any>[] = [];
          ws.getRows(2, ws.rowCount)?.forEach((row, ri) => {
            if (row.cellCount > 0) {
              rows[ri] = {};
              for (let ci = 1; ci <= cellCount; ci++) {
                let cell = row.getCell(ci).master;

                const colKey: string = columns[ci];
                rows[ri][colKey] = cell.value;
                rows[ri][colKey + ' Note'] = (
                  cell.note && (cell.note as any).texts && (cell.note as any).texts[0] && (cell.note as any).texts[0].text
                    ? (cell.note as any).texts[0].text
                    : ''
                );

                // // TODO: cell.comments doesn't exits, cell.note only include notes on cell.tnote.texts[0].text if using comments instead cell.note.texts is an empty array
                // rows[ri][colKey + ' Chat'] = cell.comments.map(comment => { return { user:comment.author, date:comment.created, message:comment.value }; });
              }
            }
          });

          workbook.push({
            name: ws.name,
            columns,
            rows,
          });
        });
    }).then(() => calculate());
  }
}

function calculate() {
  const rows: any[] = [];
  const relevantKey: string = '#' // Ticket-Präfix

  workbook.filter(ws => ws.name === 'Mai').forEach(worksheet => {
    // get the rows with relevant data
    const relevantRows: Record<string, any>[] = worksheet.rows.filter((x: any) => x.Metadaten && x.Metadaten.includes(relevantKey));

    if (relevantRows.length > 0) {
      const ticketRows: any[] = [];
      relevantRows.forEach(x => {
        const tempRows: any[] = [];
        // split rows with several tickets
        for (let meta: string = x.Metadaten; meta.includes(relevantKey);) {
          let ticket = meta.substr(meta.indexOf(relevantKey), 6);
          meta = meta.replace(ticket, '');

          tempRows.push({
            Datum: new Date(x.Datum),
            Ticket: ticket,
            Metadaten: x.Metadaten,
            Taetigkeiten: x.Taetigkeiten,
            ['Zeitaufwand (Dokumentiert)']: x['Zeitaufwand in Std.']
          });
        }

        // recalc the time for several tickets in the same relevantRows
        ticketRows.push(...tempRows.map((row: any) => {
          return {
            ...row,
            Tickets: tempRows.map(tmp => tmp.Ticket).filter(x => x !== row.Ticket),
            ['Zeitaufwand (Ticket-Anteil)']: row['Zeitaufwand (Dokumentiert)'] / tempRows.length
          }
        }));
      });

      rows.push(...ticketRows);
    }
  });

  // https://stackoverflow.com/questions/14446511/most-efficient-method-to-groupby-on-an-array-of-objects
  var groupBy = (xs: any, key: string) => {
    return xs.reduce((rv: any, x: any) => {
      (rv[x[key]] = rv[x[key]] || []).push(x);
      return rv;
    }, {});
  };

  // group by tickets
  const tickets = groupBy(rows, 'Ticket');

  // sum time needed for each ticket
  Object.keys(tickets).forEach((ticket: any) => {
    let sum: number = 0;

    Object.values(tickets[ticket]).forEach((x: any) => {
      sum += x['Zeitaufwand (Ticket-Anteil)'];
    });

    console.log(ticket, sum.toFixed(2));
    console.log(tickets[ticket].sort((a: any, b: any) => a.Datum.valueOf() - b.Datum.valueOf()).map((x: any) => {
      return `${formateDate(x.Datum)} [${x['Zeitaufwand (Ticket-Anteil)'].toFixed(2)}]${x.Tickets.length > 0 ? ` (Parallel: ${x.Tickets.join(', ')})` : ''}:\n${x.Taetigkeiten.split('\n').map((m: any) => m.trim().startsWith('-') ? '-' + m : '- ' + m).join('\n')}`;
    }).join('\n\n'));
  });
}

function formateDate(date: Date): string {
  return `${`00${date.getDate()}`.slice(-2)}.${`00${date.getMonth()+1}`.slice(-2)}.${date.getFullYear()}`;
}

function App() {
  return (
    <div className="App">
      <input id='worklog' type='file' onChange={read} />
    </div>
  );
}

export default App;
