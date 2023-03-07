function main(workbook: ExcelScript.Workbook) {
  let SHEET = workbook.getActiveWorksheet();

  var RANGE = 40;
  var SKU = "E";

  // check for duplicate SKUs
  var completed = new Set();
  var multiRows = {};

  for (let i = 2; i <= RANGE; i++) {

        let skuCell = SHEET.getRange(SKU + i);
        let skuValue = skuCell.getValue();

        if (completed.has(skuValue)) {
          multiRows[skuValue] = [];
        }
        else {
          completed.add(skuValue);
        }

  }

  for (let i = 2; i <= RANGE; i++) {

    let skuCell = SHEET.getRange(SKU + i);
    let skuValue = skuCell.getValue();

        // init cells
        let quantity = SHEET.getRange("J" + i);
        let price = SHEET.getRange("K" + i); 
        let unit = SHEET.getRange("L" + i); 
        let unit2 = SHEET.getRange("M" + i); 

        let weight = SHEET.getRange("O" + i);
        let unitW = SHEET.getRange("P" + i);

      // assign cell's value
      let value1 = quantity.getValue();
      let value2 = price.getValue();
      let value3 = weight.getValue();



    if (!(skuValue in multiRows)) {

          // calculate
          let calc1 = value2 / value1;
          let calc2 = calc1.toFixed(2);
          let calc3 = value3 / value1;

          // place calculated values in cells
          unit.setValue(calc1);
          unit2.setValue(calc2);
          unitW.setValue(calc3);

    }

    else {
      multiRows[skuValue].push(parseFloat(value1));
      multiRows[skuValue].push(parseFloat(value2));
    }
  }

  for (let row in multiRows) {

    let value1 = 0;
    let value2 = 0;
    let value3 = 0; // initialize new value to 0

        for (let i = 0; i < multiRows[row].length; i++) {

              if (i % 2 === 0) { // check if i is divisible by 3 (new value is at index 2)
                value1 += multiRows[row][i];

              } else {
                value2 += multiRows[row][i];

              }

        }

    multiRows[row] = [value1, value2]; // update the row with the new values
  }

  console.log(multiRows);

}
