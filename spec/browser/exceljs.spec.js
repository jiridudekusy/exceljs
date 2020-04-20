/* global ExcelJS */
// ExcelJS is a global injected by `./dist/exceljs.js` during jasmine's setup

'use strict';

const Stream = require('stream');

function unexpectedError(done) {
  return function(error) {
    // eslint-disable-next-line no-console
    console.error('Error Caught', error.message, error.stack);
    expect(true).toEqual(false);
    done();
  };
}

describe('ExcelJS', () => {
  it('should read and write xlsx via binary buffer', done => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = 7;

    wb.xlsx
      .writeBuffer()
      .then(buffer => {
        const wb2 = new ExcelJS.Workbook();
        return wb2.xlsx.load(buffer).then(() => {
          const ws2 = wb2.getWorksheet('blort');
          expect(ws2).toBeTruthy();

          expect(ws2.getCell('A1').value).toEqual('Hello, World!');
          expect(ws2.getCell('A2').value).toEqual(7);
          done();
        });
      })
      .catch(error => {
        throw error;
      })
      .catch(unexpectedError(done));
  });
  it('should  write xlsx into stream', async () => {
    const output = new Stream.Writable();
    const res = [];
    output._write = (chunk, encoding, next) => {
      res.push(chunk);
      next();
    };

    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({stream: output});
    const ws = wb.addWorksheet('blort');
    ws.addRow(['Hello, World!']).commit();
    ws.addRow([7]).commit();
    ws.commit();
    await wb.commit();

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(Buffer.concat(res));
    const ws2 = wb2.getWorksheet('blort');
    expect(ws2).toBeTruthy();

    expect(ws2.getCell('A1').value).toEqual('Hello, World!');
    expect(ws2.getCell('A2').value).toEqual(7);
  });
  it('should read and write xlsx via base64 buffer', done => {
    const options = {
      base64: true,
    };
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = 7;

    wb.xlsx
      .writeBuffer(options)
      .then(buffer => {
        const wb2 = new ExcelJS.Workbook();
        return wb2.xlsx.load(buffer.toString('base64'), options).then(() => {
          const ws2 = wb2.getWorksheet('blort');
          expect(ws2).toBeTruthy();

          expect(ws2.getCell('A1').value).toEqual('Hello, World!');
          expect(ws2.getCell('A2').value).toEqual(7);
          done();
        });
      })
      .catch(error => {
        throw error;
      })
      .catch(unexpectedError(done));
  });
  it('should write csv via buffer', done => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('B1').value = 'What time is it?';
    ws.getCell('A2').value = 7;
    ws.getCell('B2').value = '12pm';

    wb.csv
      .writeBuffer()
      .then(buffer => {
        expect(buffer.toString()).toEqual(
          '"Hello, World!",What time is it?\n7,12pm'
        );
        done();
      })
      .catch(error => {
        throw error;
      })
      .catch(unexpectedError(done));
  });
});
