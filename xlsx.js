const ExcelJS = require('exceljs');
const Range = require('exceljs/lib/doc/range')

const workbook = new ExcelJS.Workbook();

const worksheet = workbook.addWorksheet('Sheet1');

function addHead(ws, ref) {
    const range = new Range(ref);
    ['number', 'address', 'privatekey'].forEach(
        (item, index) => {
            ws.getCell(range.top, range.left + index).value = item;
        }
    );
}

function addRowContent(ws, ref, line, address, privatekey) {
    const range = new Range(ref);

    ws.getCell(range.top + line, range.left ).value = line;
    ws.getCell(range.top + line, range.left + 1).value = address;
    ws.getCell(range.top + line, range.left + 2).value = privatekey;
}

addHead(worksheet,"A1");
addRowContent(worksheet,"A1",1, '0x845EF9dB59b914536b1Fe884D11913EbB48bd07F', '0xd411fae188159ea3064cd7be3bc63eb604b73a23d4b18e67ac342c7bfbc8dc3c');
workbook.xlsx
    .writeFile("address.xls")
    .then(() => {
        console.log('Done.');
    })
    .catch(error => {
        console.log(error.message);
    });
