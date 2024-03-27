/**
 * Fix border SPPD Depan worksheet
 *
 * @param {ExcelJS.Worksheet} worksheet
 */
function styleSppdDepanWorksheet(worksheet) {
    worksheet.getCell('E10').border = {
        top: {style:'thin', color: {argb:'00000000'}},
        right: {style:'thin', color: {argb:'00000000'}}
    }
    worksheet.getCell('E18').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('E19').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('E20').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('E21').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('E28').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('E29').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('A22').border = {right: {style:'thin', color: {argb:'00000000'}}}

    worksheet.getCell('F15').border = {left: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F16').border = {left: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F18').border = {left: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F20').border = {left: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F22').border = {left: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F23').border = {left: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F25').border = {left: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F26').border = {left: {style:'thin', color: {argb:'00000000'}}}

    worksheet.getCell('E16').border = {bottom: {style:'thin', color: {argb:'FF000000'}}}
    worksheet.getCell('E18').border = {bottom: {style:'thin', color: {argb:'FF000000'}}}
    worksheet.getCell('E20').border = {bottom: {style:'thin', color: {argb:'FF000000'}}}
    worksheet.getCell('E15').border = {bottom: {style:'thin', color: {argb:'FFFFFFFF'}}}
    worksheet.getCell('E22').border = {bottom: {style:'thin', color: {argb:'FFFFFFFF'}}}
    worksheet.getCell('E25').border = {bottom: {style:'thin', color: {argb:'FFFFFFFF'}}}

    worksheet.getCell('E24').border = {
        top: {style:'thin', color: {argb:'00000000'}},
        right: {style:'thin', color: {argb:'00000000'}}
    }
    worksheet.getCell('E27').border = {
        top: {style:'thin', color: {argb:'00000000'}},
        right: {style:'thin', color: {argb:'00000000'}}
    }
}

/**
 * Fix border SPPD Belakang worksheet
 *
 * @param {ExcelJS.Worksheet} worksheet
 */
function styleSppdBelakangWorksheet(worksheet) {
    worksheet.getCell('F12').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F15').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F19').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F25').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F29').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F35').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F39').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F45').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F49').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F56').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('F61').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('D29').border = {right: {style:'thin', color: {argb:'00000000'}}}
    worksheet.getCell('D39').border = {right: {style:'thin', color: {argb:'00000000'}}}
}


export { styleSppdDepanWorksheet, styleSppdBelakangWorksheet }
