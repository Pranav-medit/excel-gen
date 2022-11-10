const XLSX = require('xlsx')
const fs = require('fs');
const { uniqueNamesGenerator, adjectives, colors, animals } = require('unique-names-generator');
const { generatePan, generateRandomPan, generateSerialPan } = require('./pan-generator');
const { LoremIpsum } = require("lorem-ipsum");
// const LoremIpsum = require("lorem-ipsum").LoremIpsum;

const lorem = new LoremIpsum({
  sentencesPerParagraph: {
    max: 8,
    min: 4
  },
  wordsPerSentence: {
    max: 16,
    min: 4
  }
});

fs.readFile('excel_headers.txt', 'utf8', function(err, data) {
    // cc = new a();
    let arr  =  data.trim().replaceAll('\r','').split('\n');
    arr  = arr.map((val)=> separateCharactersUponUppercase(val, true))
    const obj = arr.reduce((prev,curr)=> {
        return {...prev,[curr]:''}
    },{})
    writeExcelSheet(obj);
});
function writeExcelSheet(data){ 
    const ws = XLSX.utils.json_to_sheet([data])
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Responses')
    XLSX.writeFile(wb, 'PurchaseSale.xlsx')
}
function capitalizeFirstLetter(strToCapitalize) {
    return strToCapitalize[0].toUpperCase() + strToCapitalize.slice(1);
}
function separateCharactersUponUppercase(strInput, firstCap) {
    if (firstCap) {
        return capitalizeFirstLetter(strInput.split(/(?=[A-Z])/).join(" "));
    }
    return strInput.split(/(?=[A-Z])/).join(" ");
}


const randomData = {
    branchCode : getRandomFromArray(["001", 106, 111, 1188, 123, 888, 932, "ABR", "BLR", "BSG", "CHN", "HO", "KLK", "KRG", "NEWB", "PRT", "SLN", "SLS", "TVL", "TVM", "VLR"]),
    transactionType: getRandomFromArray(["CHARGE", "PAYMENT", "BRANCH_MERGER", "ACCOUNT_TRANSFER", "REFUND", "REBATE"]),
    transactionId: makeid(10),
    transactionLotId:makeid(11),
    emailId: makeid(6)+'@gmail.com',
    pan: generateRandomPan(),
    gstin: makeid(10),
    locationCode :getRandomFromArray(["ajk", 108, 8808, 108, 66300, "das", "sg", "das", "kar"]),
    valueDate : getRandomDate(),
    dueDate : getRandomDate(),
    description : lorem.generateWords(4),
    purchaseOrderNumber : getRandomNum(12),
    purchaseOrderDate: getRandomDate(),
    purchaseOrderDescription: lorem.generateWords(10),  
    amount: getRandomNum(5),
    currency: getRandomFromArray(['INR','USD']),
    surcharge: getRandomNum(2),
    firstName: randomName(),
    middleName: getRandomNum(2),
    lastName: makeid(2),
    address1:  makeid(3)+lorem.generateWords(3),
    address2: makeid(2)+lorem.generateWords(4),
    address3: makeid(4)+lorem.generateWords(2),
    cityCode: "212",
    districtCode: "212",
    stateCode: "96",
    phone1: getRandomNum(10), 
    phone2: getRandomNum(10),
    tdsSection: makeid(8),
    linkedInvoiceNumber: getRandomNum(5),
    reference: makeid(20),
    itemAccountCode: getRandomNum(9),
    itemDescription: lorem.generateSentences(1), 
    itemQuantity: getRandomNum(2),
    itemRate: getRandomNum(3),
    itemAmount: getRandomNum(4) ,
    itemSurcharge: getRandomNum(4),

}
console.log(randomData)
function makeid(length) {
    var result           = '';
    var characters       = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    var charactersLength = characters.length;
    for ( var i = 0; i < length; i++ ) {
        result += characters.charAt(Math.floor(Math.random() * charactersLength));
    }
    return result;
}
function getRandomNum(length) {

    return Math.floor(Math.pow(10, length-1) + Math.random() * 9 * Math.pow(10, length-1));
    
    }

function randomDate(start, end) {
    const date =  new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
    return date.toISOString().slice(0, 10);
}
function getRandomDate(){
    return randomDate(new Date(2012, 0, 1), new Date());
}
const randomName = () => uniqueNamesGenerator({ dictionaries: [adjectives, colors, animals] }); // big_red_donkey


function getRandomFromArray(arr){
    return arr[Math.floor(Math.random()*arr.length)]
}