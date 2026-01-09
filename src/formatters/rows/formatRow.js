import formatCell from '../cells/formatCell';

export default (row, index, totalRows, rowHeight, headerRowHeight) => {
  // To ensure the row number starts as in excel.
  const rowIndex = index + 1;
  const rowCells = row
    .map((cell, cellIndex) => formatCell(cell, cellIndex, rowIndex, index === 0, index === totalRows - 1, cellIndex === 0, cellIndex === row.length - 1))
    .join('');

  var heightAttr = "";
  if (rowHeight) {
    heightAttr = `ht="${rowHeight}" customHeight="1" `;
  }

  row.forEach(cell => {
    if (cell.isHeader) {
      const h = headerRowHeight || 25;
      heightAttr = `ht="${h}" customHeight="1" `;
    }
  });
  return `<row ${heightAttr}r="${rowIndex}">${rowCells}</row>`;

};
