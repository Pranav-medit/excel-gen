const XLSX = require('xlsx')
const fs = require('fs');
const { uniqueNamesGenerator, adjectives, colors, animals } = require('unique-names-generator');
const { generatePan, generateRandomPan, generateSerialPan } = require('./pan-generator');
const { LoremIpsum } = require("lorem-ipsum");
// const LoremIpsum = require("lorem-ipsum").LoremIpsum;
const customerID=[ '0000000000055462', '0000000000076535', '10101010', '101123456', '10533', '10624', '106467', '111111111', '1111122222', '1111133333', '1119000002', '1119000003', '1119000004', '1119000005', '1119000006', '1119000007', '1119000008', '1119000011', '1119000012', '1119000013', '1119000014', '1119000015', '1119000016', '1119000017', '1119000018', '1119000019', '1119000020', '1119000021', '1119000022', '1119000023', '1119000024', '1119000025', '1119000026', '1119000027', '1119000028', '1119000029', '1119000030', '1119000031', '1119000032', '1119000033', '1119000034', '1119000035', '1119000036', '1119000037', '1119000038', '1119000039', '1119000040', '1119000041', '1119000042', '1119000043', '1119000044', '1119000045', '1119000046', '1119000047', '1119000048', '1119000049', '1119000050', '1119000051', '1119000052', '1119000053', '1119000054', '1119000055', '1119000056', '1119000057', '1119000058', '1119000059', '1119000060', '1119000061', '1119000062', '1119000063', '1119000064', '1119000067', '1119000068', '1119000069', '1119000070', '1119000071', '1119000072', '1119000073', '1119000074', '1119000075', '1119000076', '1119000077', '1119000078', '1119000079', '1119000080', '1119000081', '1119000082', '1119000083', '1119000084', '1119000086', '1119000087', '1119000088', '1119000089', '1119000090', '1119000091', '1119000092', '1119000093', '1119000094', '1119000095', '1119000096', '1119000097', '1119000098', '1119000099', '1119000100', '1119000101', '1119000102', '1119000103', '1119000104', '1119000105', '1119000106', '1119000107', '1119000108', '1119000109', '1119000110', '1119000111', '1119000112', '1119000113', '1119000114', '1119000115', '1119000116', '1119000117', '1119000118', '1119000119', '1119000120', '1119000121', '1119000122', '1119000123', '1119000124', '1119000125', '1119000126', '1119000127', '1119000128', '1119000129', '1119000130', '1119000131', '1119000132', '1119000133', '1119000134', '1119000135', '1119000136', '1119000137', '1119000138', '1119000139', '1119000140', '1119000141', '1119000142', '1119000143', '1119000144', '1119000145', '1119000146', '1119000147', '1119000148', '1119000149', '1119000150', '1119000151', '1119000152', '1119000153', '1119000154', '1119000155', '1119000156', '1119000157', '1119000158', '1119000159', '1119000160', '1119000161', '1119000162', '1119000163', '1119000164', '1119000165', '1119000166', '1119000167', '1119000168', '1119000169', '1119000170', '1119000171', '1119000172', '1119000173', '1119000174', '1119000175', '1119000176', '1119000177', '1119000178', '1119000179', '1119000180', '1119000181', '1119000182', '1119000183', '1119000184', '1119000185', '1119000186', '1119000187', '1119000188', '1119000189', '1119000190', '1119000191', '1119000192', '1119000193', '1119000194', '1119000195', '1119000196', '1119000197', '1119000198', '1119000199', '1119000200', '1119000201', '1119000202', '1119000203', '1119000204', '1119000205', '1119000206', '1119000207', '1119000208', '1119000209', '1119000210', '1119000211', '1119000235', '1119000236', '1119000237', '1119000238', '1119000239', '1119000240', '1119000241', '1119000242', '1119000243', '1119000244', '1119000245', '1119000246', '1119000247', '1119000248', '1119000249', '1119000250', '1119000251', '1119000252', '1119000253', '1119000254', '1119000255', '1119000256', '1119000257', '1119000258', '1119000259', '1119000260', '1119000261', '1119000262', '1119000263', '1119000264', '1119000265', '1119000266', '1119000267', '1119000268', '1119000269', '1119000270', '1119000271', '1119000272', '1119000273', '1119000274', '1119000275', '1119000277', '1119000278', '1119000279', '1119000280', '1119000284', '1119000285', '1119000286', '1119000287', '1119000288', '1119000289', '1119000290', '1119000291', '1119000292', '1119000293', '1119000294', '1119000295', '1119000296', '1119000297', '1119000298', '1119000299', '1119000300', '1119000301', '1119000302', '1119000303', '1119000304', '1119000305', '1119000306', '1119000307', '1119000308', '1119000309', '1119000310', '1119000311', '1119000312', '1119000313', '1119000314', '1119000315', '1119000316', '1119000317', '1119000318', '1119000319', '1119000320', '1119000321', '1119000322', '1119000323', '1119000324', '1119000325', '1119000326', '1119000327', '1119000329', '1119000330', '1119000331', '1119000332', '1119000333', '1119000334', '1119000335', '1119000336', '1119000337', '1119000338', '1119000339', '1119000340', '1119000341', '1119000342', '1119000343', '1119000344', '1119000345', '1119000346', '1119000347', '1119000348', '1119000349', '1119000350', '1119000351', '1119000352', '1119000353', '1119000354', '1119000355', '1119000356', '1119000357', '1119000358', '1119000359', '1119000360', '11asd345', '12', '123', '1234', '1234123412341234', '12345', '123456', '123456741258', '12345678', '123456789', '127895634', '132', '136736', '162344', '20567', '20615', '20626', '20644', '213', '233193', '334455', '41858', '4455663', '44961', '4810', '560021457896', '6219', '6811', '6855', '7014', '754511', '7894', 'aaa45asa45', 'aaa?*asa45', 'ALPHA01', 'chitra', 'T8JQIAV640K75RUOC8R9', 'TestCust00001', 'TestCust00002', 'TestCust00004', 'TestCust00006', 'ZSL2735629']

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

function writeExcelSheet(data){ 
    const ws = XLSX.utils.json_to_sheet(data)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Responses')
    XLSX.writeFile(wb, 'output/excel-sheets/PurchaseSaleInvoice.xlsx')
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


const randomData=  () => {
    const itemQuantity = getRandomNum(2)
    const itemRate =  getRandomNum(3)
    const itemAmount = itemQuantity * itemRate;
   return  {
    branchCode : getRandomFromArray(["HO"]),
    transactionType: getRandomFromArray(["PURCHASE_INVOICE"]),
    transactionId: makeid(10),
    transactionLotId:makeid(11),
    emailId: makeid(6)+'@gmail.com',
    pan: generateRandomPan(),
    gstin: makeid(10),
    locationCode :getRandomFromArray(["sg"]),
    valueDate : getRandomDate(),
    dueDate : getRandomDate(),
    description : lorem.generateWords(4),
    purchaseOrderNumber : getRandomNum(12),
    purchaseOrderDate: getRandomDate(),
    purchaseOrderDescription: lorem.generateWords(10),  
    amount: getRandomNum(5),
    currency: getRandomFromArray(['INR']),
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
    itemAccountCode: "PuSaTdsPay",
    itemDescription: lorem.generateSentences(1), 
    itemQuantity: itemQuantity,
    itemRate:itemRate,
    itemAmount: itemAmount ,
    itemSurcharge: getRandomNum(4),
    countryCode: "USA",
    pinCode: getRandomNum(6),
    taxMode: getRandomFromArray(["EXCLUSIVE","INCLUSIVE"]),
    terms: getRandomFromArray(["NET_15","NET_30","NET_60","NET_90"]),
    itemTaxRate: Math.random(),
    customerId:getRandomFromArray(['1119000002'])
}}
// for(const i in randomData){
//     console.log(i,randomData[i])
// }
// let arr =[];
let arr = []
let max = 10
for(let i=0;i<max;i++){
    arr.push(randomData())
}
// console.log(arr)
writeExcelSheet(arr);
// fs.readFile('excel_headers.txt', 'utf8', function(err, data) {
//     // cc = new a();
//     let arr  =  data.trim().replaceAll('\r','').split('\n');
//     arr  = arr.map((val)=> separateCharactersUponUppercase(val, true))
//     const obj = arr.reduce((prev,curr)=> {
//         return {...prev,[curr]:''}
//     },{})
//     writeExcelSheet(obj);
// });

// bc=["001", 106, 111, 1188, 123, 888, 932, "ABR", "BLR", "BSG", "CHN", "HO", "KLK", "KRG", "NEWB", "PRT", "SLN", "SLS", "TVL", "TVM", "VLR"]'INR',