
const XLSX = require('xlsx')
const input = [
  "accountId",
  "transactionId",
  "transactionName",
  "valueDate",
  "transactionDate",
  "amount1",
  "amount2",
  "amount3",
  "part1",
  "part2",
  "part3",
  "part4",
  "part5",
  "description",
  "payeeAccountId",
  "userId",
  "param1",
  "param2",
  "instrument",
  "reference",
  "part6",
  "currencyCode",
  "partyInfo",
  "bankAccountNumber"
]
function writeExcelSheet(data){ 
  const ws = XLSX.utils.json_to_sheet(data)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Responses')
  XLSX.writeFile(wb, 'output/excel-sheets/LoanOdDisbursement.xls')
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

const obj = input.reduce((prev,curr)=> {
          return {...prev,[separateCharactersUponUppercase(curr,true)]:''}
      },{})
    writeExcelSheet([obj]);