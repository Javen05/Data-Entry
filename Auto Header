function main(workbook: ExcelScript.Workbook) {
    let SHEET = workbook.getActiveWorksheet();

    const cells = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"];
    const values = ["SN", "Inbound AWB", "Permit_Number", "ASN Number", "SKU", "Concat(Permit+SKU)", "Description", "HS_Code", "COO", "Inbound Qty", "Permit declared value", "Unit_Price", "Unit Price to use", "Balance qty in warehouse", "Total_Weight", "Unit Weight", "Remarks"];
    
    for (let i = 0; i < cells.length; i++) {
        let a = SHEET.getRange(cells[i] + "1");
        a.setValue(values[i]);
    }
}

