import formatCell from '../cells/formatCell';

export default (row, index) => {
    // To ensure the row number starts as in excel.
  const rowIndex = index + 1;
  const rowCells = row
    .map((cell, cellIndex) => formatCell(cell, cellIndex, rowIndex))
    .join('');
  
  var heightAttr = "";
  row.forEach(cell=>{
    if (cell.isHeader) {
        heightAttr = 'ht="25" customHeight="1"';
    }
  });
  return `<row ${heightAttr} r="${rowIndex}">${rowCells}</row>`;
};