const https = require('https');

https.get('https://restcountries.com/v3.1/all', res => {

    var xl = require('excel4node');
    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('Sheet 1');

    var styleTitle = wb.createStyle({
        alignment: {
            wrapText: true,
            horizontal: 'center',
        },
        font: {
            bold: true,
            color: '#4F4F4F',
            size: 16,
        },
    });

    var styleColumns = wb.createStyle({
        font: {
            bold: true,
            color: '#808080',
            size: 12,
        },
    });

    var styleArea = wb.createStyle({
        numberFormat: '#,00.00; -#,00.00; -'
    });

    let data = [];

    res.on('data', chunk => {
        data.push(chunk);
    });

    res.on('end', () => {

        const countries = JSON.parse(Buffer.concat(data).toString());

        ws.cell(1, 1, 1, 4, true).string('Countries List').style(styleTitle);

        ws.cell(2, 1)
            .string('Name')
            .style(styleColumns);

        ws.cell(2, 2)
            .string('Capital')
            .style(styleColumns);

        ws.cell(2, 3)
            .string('Area')
            .style(styleColumns);

        ws.cell(2, 4)
            .string('Currencies')
            .style(styleColumns);

        let row = 3;

        countries.forEach(country => {
            ws.cell(row, 1).string(country.name.common);

            if (country.capital) {
                ws.cell(row, 2).string(country.capital);
            } else {
                ws.cell(row, 2).string("-");
            }

            if (country.area) {
                ws.cell(row, 3).number(country.area).style(styleArea);
            } else {
                ws.cell(row, 3).string("-");
            }

            let currencies = [];
            if (country.currencies) {
                Object.keys(country.currencies).forEach(function (currency) {
                    currencies.push(currency);
                });
                ws.cell(row, 4).string(String(currencies)).style({ italics: false });
            } else {
                ws.cell(row, 4).string("-");
            }
            console.log(`${country.name.common} cell added!`);
            row++;
        });
        wb.write('Excel.xlsx');
    });
}).on('error', err => {
    console.log('Error: ', err.message);
});

