const Web3 = require('web3')
const ExcelJS = require('exceljs');
const Range = require('exceljs/lib/doc/range')

const web3 = new Web3(new Web3.providers.HttpProvider("http://127.0.0.1:7545"));

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

function createAccounts(){
    for(var i=0; i < 100000; i++) {
        var wallet = web3.eth.accounts.create();
/*        console.log("address:" + wallet.address);
        console.log("privateKey:" + wallet.privateKey);*/
        addRowContent(worksheet,"A1",i+1, wallet.address, wallet.privateKey);
    }
}

createAccounts();

workbook.xlsx
    .writeFile("address.xls")
    .then(() => {
        console.log('Done.');
    })
    .catch(error => {
        console.log(error.message);
    });
