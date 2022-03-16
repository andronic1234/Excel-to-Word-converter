const XLSX = require('xlsx');
const fs = require('fs');

const parse = (filename) => {
    const eData = XLSX.readFile(filename);

    return Object.keys(eData.Sheets).map((name) => ({
        name,
        data: XLSX.utils.sheet_to_json(eData.Sheets[name], { header: 0 }),
    }));
};

const ExcelFiles = fs.readdirSync('./Excel/').filter(file => file.endsWith('.xlsx'));


    parse(`./Excel/${ExcelFiles}`).forEach((element) => {
        console.log(element.data)
        const finished = (error) => {
            if(error){
                console.log(error);
                return;
            }
        }
        const JSONstorer = JSON.stringify(element.data, null, 2)
        fs.writeFile('storer.json', JSONstorer, finished)

    });

