import XlsxPopulate, { RichText } from 'xlsx-populate';

export function createXlsx() {
    console.log('11111');
    XlsxPopulate.fromBlankAsync().then(workbook => {
        console.log('들어옴');
        const rtext = new RichText();
        rtext.add('aa1a').add(`hello\ndddddd`, { italic: true, bold: true })
             .add('world!', { fontColor: 'FF0000' });
        workbook.sheet('Sheet1').cell('A1').value(rtext);
                // .add('hello ', { italic: true, bold: true })
                // .add('world!', { fontColor: 'FF0000' });

        return workbook.outputAsync()
                       .then(function (blob) {
                           if (window.navigator && window.navigator.msSaveOrOpenBlob) {
                               // If IE, you must uses a different method.
                               window.navigator.msSaveOrOpenBlob(blob, '응시자_일괄등록_양식.xlsx');
                           } else {
                               const url = window.URL.createObjectURL(blob);
                               const a = document.createElement('a');
                               document.body.appendChild(a);
                               a.href = url;
                               a.download = '응시자_일괄등록_양식.xlsx';
                               a.click();
                               window.URL.revokeObjectURL(url);
                               document.body.removeChild(a);
                           }
                       });
    });
}
