/**
 * Created by samsung np on 16.06.2016.
 */
/**
 * Created by samsung np on 12.06.2016.
 */
var Excel = require('exceljs');
var fs = require('fs');
/*var loadedFile=__dirname + "/uploads/c313_2015-10_sum.xlsx";*/

/*var n=5;*/
/*var finalizedRow =[10,20,30,40,50,60,70,80,90,100,110,120,130,140,150,160,170,180]*/



/*
function presenceInputer(house, book){
    var sheet2 = book.addWorksheet('лист1');
    var worksheet = book.getWorksheet('лист1');
    fse.removeSync(__dirname + '/db.xlsx');
    console.log("книга удалена")
    book.xlsx.writeFile(__dirname + '/db.xlsx')
        .then(function() {
            // done
        });
}
function voidinputer(house){
    fse.removeSync(__dirname + '/db.xlsx');
    var workbook = new Excel.Workbook();
    var sheet = workbook.addWorksheet(1);
    var worksheet = workbook.getWorksheet(1);
    worksheet.getCell('A1').value = house.company;
    worksheet.getCell('B1').value = house.adress;
    worksheet.getCell('C1').value = house.year;
    worksheet.getCell('D1').value = house.month;
    worksheet.getCell('E1').value = house.owner;
    worksheet.getCell('F1').value = house.ownersCount;
    worksheet.getCell('G1').value = house.square;
    worksheet.getCell('H1').value = house.lgotSquaree;
    worksheet.getCell('I1').value = house.techOmain;
    worksheet.getCell('J1').value = house.techOOther;
    worksheet.getCell('K1').value = house.naemSoc;
    worksheet.getCell('L1').value = house.naemKomm;
    worksheet.getCell('M1').value = house.naemBezDot;
    worksheet.getCell('N1').value = house.heat;
    worksheet.getCell('O1').value = house.capRepair;
    worksheet.getCell('P1').value = house.drainage;
    worksheet.getCell('Q1').value = house.waterSupply;
    worksheet.getCell('R1').value = house.eEnergy;
    worksheet.getCell('S1').value = house.compensation;
    worksheet.getCell('T1').value = house.changeSum;

    workbook.xlsx.writeFile(__dirname + '/db.xlsx')
        .then(function() {
            // done
        });
};
*/




var sumreader = function (Lmonth) {
    var almostparsed = Lmonth.split('_');
    var fullyparsed = almostparsed[1].split('-');
    var pathToFile = __dirname + '/uploads/' + Lmonth;
    tryer(fullyparsed[0], fullyparsed[1], pathToFile);
    console.log(fullyparsed[0], fullyparsed[1]);
};

var tryer = function (year, month,loadedFile) {
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(loadedFile)
        .then(function () {
            var worksheet = workbook.getWorksheet('Лист1');
            var row = worksheet.getRow(worksheet.lastRow.number-3);
            console.log(loadedFile);
            console.log(row.values);
            var controlCompany=worksheet.getCell('A6').value.split(':').splice(1, 1).join();
            var finalizedRow=row.values;
            finalizedRow.push(controlCompany);
            var house = new House(year, month, finalizedRow);
            dbinputer(house)
        });
};


    function dbinputer (house){
        var workbook = new Excel.Workbook();
        workbook.xlsx.readFile(__dirname + '/db.xlsx')
            .then(function (book) {
                var worksheet = workbook.getWorksheet(1);
                var allowInput=true;
                worksheet.eachRow(function(row, rowNumber) {
                    console.log("компания из House" + house.company);
                    console.log("компания из Row" + row.values[1]);
                    if(house.company==row.values[1]&&house.year==row.values[3]&&house.month==row.values[4]){
                        allowInput=false;
                        console.log("совпадение нашлось и запись " + allowInput)}
                });
                console.log("итоговое значение разрешателя "+ allowInput);
                if(allowInput==true){
                    console.log("неужели?!");
                    inputer(house,book);
                }
            });

    }

    function House(year, month,finalizedRow){
        this.adress="0";
        this.year=year;
        this.month=month;
        this.owner=finalizedRow[1];
        this.ownersCount=finalizedRow[2];
        this.square=finalizedRow[3];
        this.lgotSquaree=finalizedRow[4];
        this.techOmain=finalizedRow[5];
        this.techOOther=finalizedRow[6];
        this.naemSoc=finalizedRow[7];
        this.naemKomm=finalizedRow[8];
        this.naemBezDot=finalizedRow[9];
        this.heat=finalizedRow[10];
        this.capRepair=finalizedRow[11];
        this.drainage=finalizedRow[12];
        this.waterSupply=finalizedRow[13];
        this.eEnergy=finalizedRow[14];
        this.compensation=finalizedRow[15];
        this.changeSum=finalizedRow[16];
        this.company=finalizedRow[17];

    }

    function inputer(house, book){
        var worksheet = book.getWorksheet(1);
        var row = worksheet.lastRow.number+1;

        console.log(" Состояние данных"+ house.square);
        worksheet.getCell('A'+row).value = house.company;
        worksheet.getCell('B'+row).value = house.adress;
        worksheet.getCell('C'+row).value = house.year;
        worksheet.getCell('D'+row).value = house.month;
        worksheet.getCell('E'+row).value = house.owner;
        worksheet.getCell('F'+row).value = house.ownersCount;
        worksheet.getCell('G'+row).value = house.square;
        worksheet.getCell('H'+row).value = house.lgotSquaree;
        worksheet.getCell('I'+row).value = house.techOmain;
        worksheet.getCell('J'+row).value = house.techOOther;
        worksheet.getCell('K'+row).value = house.naemSoc;
        worksheet.getCell('L'+row).value = house.naemKomm;
        worksheet.getCell('M'+row).value = house.naemBezDot;
        worksheet.getCell('N'+row).value = house.heat;
        worksheet.getCell('O'+row).value = house.capRepair;
        worksheet.getCell('P'+row).value = house.drainage;
        worksheet.getCell('Q'+row).value = house.waterSupply;
        worksheet.getCell('R'+row).value = house.eEnergy;
        worksheet.getCell('S'+row).value = house.compensation;
        worksheet.getCell('T'+row).value = house.changeSum;

        console.log(worksheet.getCell('A'+row).value);
        book.xlsx.writeFile(__dirname + '/db.xlsx')
            .then(function() {
                // done
            });
    }

/*

    var number="data"+n;

    var useMonth=Lmonth+".xlsx";
    console.log(useMonth)
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(useMonth)
        .then(function () {
            var worksheet = workbook.getWorksheet('My Sheet');
            var row = worksheet.lastRow;
            var finalizedRow=row.values;

            if (n==1)            {

                console.log(fs.writeFileSync('t.json', JSON.stringify(data, null, "\t")));
                fs.writeFileSync('t.json', JSON.stringify(data, null, "\t"));
            } else {};
            n = fs.readFileSync('t.json',"utf-8")
            data=JSON.parse(n);
            data[number].adress=workbook.getWorksheet('My Sheet').getCell('C4').value;
            data[number].year=workbook.getWorksheet('My Sheet').getCell('B4').value;
            data[number].month=workbook.getWorksheet('My Sheet').getCell('A4').value;
            data[number].owner=finalizedRow[1];
            data[number].ownersCount=finalizedRow[2];
            data[number].square=finalizedRow[3];
            data[number].lgotSquare=finalizedRow[4];
            data[number].techOmain=finalizedRow[5];
            data[number].techOOther=finalizedRow[6];
            data[number].naemSoc=finalizedRow[7];
            data[number].naemKomm=finalizedRow[8];
            data[number].naemBezDot=finalizedRow[9];
            data[number].heat=finalizedRow[10];
            data[number].capRepair=finalizedRow[11];
            data[number].drainage=finalizedRow[12];
            data[number].waterSupply=finalizedRow[13];
            data[number].eEnergy=finalizedRow[14];
            data[number].compensation=finalizedRow[15];

            // use workbook
            console.log( data[number].ownersCount);
            fs.writeFileSync('t.json', JSON.stringify(data, null, "\t"));

        })*/



exports.sumreader=sumreader;

