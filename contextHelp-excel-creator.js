
const XLSX = require('xlsx');
console.log("Running")
const inputData = [
  "repayment-method","amount-range","maturity-date-reference","normal-interest-rate","normal-interest-rate-min","penal-interest-rate","penal-interest-rate-min","post-maturity-interest-rate-applicable","post-maturity-normal-interest-rate","post-maturity-demand-interval","normal-interest-limit-allowed","normal-interest-rate-basis","repayment-at-start-of-interval","interest-table-lookup","interest-table-code","interest-computation-method","interest-accrual-calculation","moratorium-type","moratorium-period","moratorium-normal-interest-rate-applicable","moratorium-normal-interest-rate","moratorium-interest-accrual-calculation","loan-purpose-code","broken-period-interest-demand-mode","interest-adjustment-mode","max-first-repayment-interval","demand-date-mode","demand-day-pattern","demand-cut-off-day","guarantor","minimum-repayment-percent","minimum-repayment-amount","multiple-disbursements-allowed","disbursement-effect","stop-moratorium-on-full-disbursement","revolving-principal-allowed","advance-payment-allowed","partial-payment-allowed","prepayment-allowed","preclosure-allowed","minimum-days-of-interest-for-preclosure","internal-preclosure-allowed","time-unit-override-allowed","repayment-method-override-allowed","prepay-interest-allowed","late-fee-as-penalty-applicable","penal-interest-for-paybyday-period","demand-pay-by-days","penal-interest-booking-mode","penal-interest-demand-mode","penal-interest-calculation","penal-interest-booking-basis","compound-interest-calculation","compounding-interval","security-deposit-deduction-type","maximum-advance-installments","prepay-to-security-deposit","security-deposit-deduction-from-customer-only","auto-close-with-security-deposit","security-deposit-due-applicable","auto-repay-scheduled","auto-repay-on-demand","holiday-handling-mode","customer-limit-checking-mode","collateral","tolerance","demand-tolerance","demand-offset-sequence","npa-demand-offset-sequence","flow-collection-sequence","npa-days","npa-unmarking-only-on-zero-dpd","carry-over-linked-account-npa","max-inactive-days","write-off-days","provisioning-applicable","provisioning-table-code","provisioning-table-code-alt","minimum-installment-amount","installment-computation-method","installment-rounding-unit","interest-computation-fraction-digits","interest-fraction-digits","interest-rate-fraction-digits","fee-charge-fraction-digits","fee-fraction-digits","demand-rounding-mode","cumulative-interest-computation","carry-over-booked-not-due-interest","extra-day-interest-on-first-disbursement-date"
];
const columns = ["Serial Number","Help Code","Help Title","Help Description"];
function writeExcelSheet(data){ 
  const style = {
    font: { color: { rgb: "FFFF0000" }, sz: 14, bold: true },
    fill: { fgColor: { rgb: "FF808080" } }
  }
  const ws = XLSX.utils.json_to_sheet(data,{ header: [], origin: -1, cellStyles: [null, style] });
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws)
  XLSX.writeFile(wb, 'output/excel-sheets/ContextHelp1.xls')
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
function separateCharUponX(string,x){
  const str =  string.trim().split(x).join(" ").trim();
  return capitalizeFirstLetter(str);
}

const obj = inputData.map( (val,index) => ({
  [ columns[0] ] :index+1,
  [ columns[1] ] : val,
  [ columns[2] ] : separateCharUponX(val,'-'),
  [ columns[3] ] : ""
}))
writeExcelSheet(obj);