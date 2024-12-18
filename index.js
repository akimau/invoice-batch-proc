const ExcelJS = require('exceljs');

const batchXLSX = "./batches/ttxls.xlsx"

// column names
const batchRefColName = "Jia Advance #";
const repaymentDateColName = "Repayment Date";
const paymentRefColName = "Ref No.";
const repaymentAmountColName = "Repayment Amount";
const totalAmountDueColName = "Total Amount Due";

// derived keys
const balanceKey= "Balance";
const repaymentsKey = "Repayments";
const totalRepaymentAmountKey = "Total Repayment Amount";
const unappliedRepaymentsKey = "Unapplied Repayments";
const unappliedOverRepaymentsKey = "Unapplied Over Repayments";
const totalUnappliedRepaymentAmountKey = "Total Unapplied Repayment Amount";
const totalUnappliedOverRepaymentAmountKey = "Total Unapplied Over Repayment Amount";
const batchCountKey = "Batch Count";


function getColsMap(worksheet) {
    const headerRow = worksheet.getRow(1); // expecting row 1 to be column headers
    let colsMap = {};
    headerRow.eachCell((cell,i) => {
        colsMap = {...colsMap, [i]: cell.value};
    });
    console.log(colsMap);
    return colsMap;
}

function parseData(worksheet, colsMap) {
    // Iterate over all rows that have values in a worksheet
    const data = [];
    worksheet.eachRow(function(row, rowNumber) {
        if(rowNumber == 1) return; // skip row 1 which is for column headers
        let rowMap = {};
        row.values.forEach((value,i) => {
            rowMap = {...rowMap, [colsMap[i]]: value};
        });
        data.push(rowMap);
    });
    return data;
}

async function loadData() {
    const workbook = await new ExcelJS.Workbook().xlsx.readFile(batchXLSX);
    const repaymentsSheet = workbook.worksheets[1];
    const batchesSheet = workbook.worksheets[2];
    const repayments  = parseData(repaymentsSheet, getColsMap(repaymentsSheet));
    const batches  = parseData(batchesSheet, getColsMap(batchesSheet));

    return {repayments, batches};
}

async function applyRepayments(data = {repayments: [], batches: []}) {
    let {repayments, batches} = data;

    if(!repayments || !batches) return [];

    let unappliedRepayments = [];

    let overRepayments = [];
    
    repayments.forEach(repayment => {
        if(!repayment[batchRefColName]) return;

        let appliedRepayment = false;

        batches.forEach(batch => {
            if(appliedRepayment) return;

            if(batch[batchRefColName] == repayment[batchRefColName]) {
                let amount = repayment[repaymentAmountColName];
                let batchBal = batch[balanceKey];
                let totalAmountDue =  batchBal == null ? batch[totalAmountDueColName] : batchBal;
                let newTotalAmountDue = totalAmountDue - amount;
                let appliedRepayments = batch[repaymentsKey] ?? [];

                var repaymentAmount = amount;

                if(newTotalAmountDue < 0) {
                    repaymentAmount = totalAmountDue;
                    batch[balanceKey] = 0;
                    let overRepayment = {
                        ...repayment, 
                        [repaymentAmountColName]: Math.abs(newTotalAmountDue)
                    };
                    overRepayments.push(overRepayment);                    
                } else {
                    batch[balanceKey] = newTotalAmountDue;
                }

                if(repaymentAmount > 0) { 
                    batch[repaymentsKey] = [...appliedRepayments, {
                        ...repayment, 
                        [repaymentAmountColName]: repaymentAmount}
                    ];
                    batch[totalRepaymentAmountKey] = (batch[totalRepaymentAmountKey] ?? 0) + repaymentAmount;
                }

                appliedRepayment = true;
            }
        });

        if(!appliedRepayment) unappliedRepayments.push(repayment);
    });

    return {overRepayments, unappliedRepayments, batches}
}

async function applyOverRepayments(data = {overRepayments: [], unappliedOverRepayments: [], batches: []}) {
    let {overRepayments : repayments, unappliedRepayments, unappliedOverRepayments, batches} = data;

    if(!repayments || !batches) return [];

    if(repayments.length === 0) return { 
        [unappliedRepaymentsKey]: unappliedRepayments,
        [unappliedOverRepaymentsKey]: unappliedOverRepayments, 
        batches,
        [batchCountKey]: batches.length,
        [totalRepaymentAmountKey] : batches.reduce(
            (p, v) => p + (v[totalRepaymentAmountKey] ?? 0), 0),
        [totalUnappliedRepaymentAmountKey] : unappliedRepayments.reduce(
            (p, v) => p + (v[repaymentAmountColName] ?? 0), 0),
        [totalUnappliedOverRepaymentAmountKey] : unappliedOverRepayments.reduce(
            (p, v) => p + (v[repaymentAmountColName] ?? 0), 0)
    }; // base case

    let overRepayments = [];

    repayments.forEach(repayment => {
        let appliedRepayment = false;

        batches.forEach(batch => {
            if(appliedRepayment) return;
            
            let batchBal = batch[balanceKey];
            let totalAmountDue =  batchBal == null ? batch[totalAmountDueColName] : batchBal;

            if(totalAmountDue > 0) {
                let amount = repayment[repaymentAmountColName];
                let newTotalAmountDue = totalAmountDue - amount;
                let appliedRepayments = batch[repaymentsKey] ?? [];

                var repaymentAmount = amount;

                if(newTotalAmountDue < 0) {
                    repaymentAmount = totalAmountDue;
                    batch[balanceKey] = 0;
                    let overRepayment = {
                        ...repayment, 
                        [repaymentAmountColName]: Math.abs(newTotalAmountDue)
                    };
                    overRepayments.push(overRepayment);
                } else {
                    batch[balanceKey] = newTotalAmountDue;
                }
                
                batch[repaymentsKey] = [
                    ...appliedRepayments, 
                    {...repayment, [repaymentAmountColName]: repaymentAmount}
                ];
                batch[totalRepaymentAmountKey] = (batch[totalRepaymentAmountKey] ?? 0) + repaymentAmount;

                appliedRepayment = true;
            }
        });
        
        if(!appliedRepayment) unappliedOverRepayments.push(repayment);
    });

    return applyOverRepayments({overRepayments, unappliedRepayments, unappliedOverRepayments, batches})
}

async function processBatch() {
    let data = await loadData(); // load repayments and batches
    data = await applyRepayments(data);
    data = await applyOverRepayments({...data, unappliedOverRepayments: []});
    console.log(JSON.stringify(data))
}

processBatch();
