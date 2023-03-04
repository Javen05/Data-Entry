<code>
function main(workbook: ExcelScript.Workbook) {
    let SHEET = workbook.getActiveWorksheet();

    var RANGE = 7;
    var SKU = "F";

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
        let cell1 = SHEET.getRange("A" + i);
        let cell2 = SHEET.getRange("B" + i);
        let cell3 = SHEET.getRange("C" + i);
        let cell4 = SHEET.getRange("D" + i);

        // assign cell's value
        let value1 = cell1.getValue();
        let value2 = cell2.getValue();

        if (!(skuValue in multiRows)) {

            // calculate
            let calc1 = value1 / value2;
            let calc2 = calc1.toFixed(2);

            // place calculated values in cells
            cell3.setValue(calc1);
            cell4.setValue(calc2);
        }

        else {
          multiRows[skuValue].push(parseFloat(value1));
          multiRows[skuValue].push(parseFloat(value2));
        }
    }

  for (let row in multiRows) {

      let value1 = 0;
      let value2 = 0;

      for (let i = 0; i < multiRows[row].length; i++) {

      if (i % 2 === 0) {
        value1 += multiRows[row][i];

      } else {
        value2 += multiRows[row][i];
      }

    }

    multiRows[row] = [value1, value2]
  }
  
  console.log(multiRows)

}
                                                </code>