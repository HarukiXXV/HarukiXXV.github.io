const { FormulaType } = require("exceljs");
const { setMaxIdleHTTPParsers } = require("http");

function exportToExcel() {
    const bag1Value = Number(document.getElementById('bag1Input').value);
    const bag2Value = Number(document.getElementById('bag2Input').value);
    const bag3Value = Number(document.getElementById('bag3Input').value);
    const dryingValue = Number(document.getElementById('dryingInput').value);
    const washingValue = Number(document.getElementById('washingInput').value);
    const paymentsValue = Number(document.getElementById('paymentsInput').value);
    const expensesValue = Number(document.getElementById('expensesInput').value);
    const cleaningCharges = Number(document.getElementById('cleaningCharges').value);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Invoice');

    sheet.addRow(['Items', 'Costs']);
    sheet.addRow(['Bag 1',`$${bag1Value}`]);
    sheet.addRow(['Bag 2',`$${bag2Value}`]);
    sheet.addRow(['Bag 3',`$${bag3Value}`]);
    const cal = ((bag1Value+bag2Value+bag3Value)*1.3)+dryingValue+washingValue+paymentsValue+expensesValue
    sheet.addRow(['Sub Total',`$${(cal.toFixed(2))}`]);
    sheet.addRow(['Cleaning Service'])
    sheet.addRow(['Tuesday',`$50`])
    sheet.addRow(['Thursday',`$50`])
    sheet.addRow(['Sunday',`$80`])
    sheet.addRow(['Additional Charges',cleaningCharges])
    const cleaningChargersTotal = 180+cleaningCharges
    sheet.addRow(['Sub Total',`$${(cleaningChargersTotal.toFixed(2))}`])
    sheet.addRow(['Total',`$${((cal+cleaningChargersTotal).toFixed(2))}`])



    const firstRow = sheet.getRow(1);
    const secondRow = sheet.getRow(5);
    const thirdRow = sheet.getRow(6);
    const forthRow = sheet.getRow(11)
    const fifthRow = sheet.getRow(12)



    firstRow.font = {
        bold: true, color: { argb: 'FFFFFF' }
    };
    firstRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '4472C4' },
    };
    secondRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '70ad47' },
    };
    secondRow.font = {
        bold: true, color: { argb: 'FFFFFF' }
    };
    thirdRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFA500' },
    };
    thirdRow.font = {
        bold: true, color: { argb: 'FFFFFF' }
    };
    forthRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '70ad47' },
    };
    forthRow.font = {
        bold: true, color: { argb: 'FFFFFF' }
    };
    fifthRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFA500' },
    };
    fifthRow.font = {
        bold: true, color: { argb: 'FFFFFF' }
    };

    const columnA = sheet.getColumn('A');
    columnA.width = 20;

    const columnB = sheet.getColumn('B');
    columnB.width = 40;

    const column = sheet.getColumn('B');

    column.eachCell((cell) => {
        cell.alignment = {
            horizontal: 'center',
            vertical: 'middle'
        };
    });

    workbook.xlsx.writeBuffer().then(function (buffer) {
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');

        link.href = url;
        link.download = 'generated_excel.xlsx';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    });
}
