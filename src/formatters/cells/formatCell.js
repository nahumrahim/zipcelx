import { validTypes, CELL_TYPE_STRING, WARNING_INVALID_TYPE } from '../../commons/constants';
import generatorStringCell from './generatorStringCell';
import generatorNumberCell from './generatorNumberCell';

export default (cell, index, rowIndex, isFirstRow, isLastRow, isFirstCol, isLastCol) => {
  if (validTypes.indexOf(cell.type) === -1) {
    console.warn(WARNING_INVALID_TYPE);
    cell.type = CELL_TYPE_STRING;
  }

  let styleIndex = 0; // Default

  if (cell.isHeader || isFirstRow) {
    styleIndex = 1; // Header
  } else {
    // Data Outline Logic
    // 2: Top-Left, 3: Top, 4: Top-Right
    // 5: Left, 6: Center, 7: Right
    // 8: Bottom-Left, 9: Bottom, 10: Bottom-Right

    // Note: Data block starts AFTER header. 
    // If header is row 0, data starts at row 1.
    // So "Top" of data block is actually row 1 (index 1).
    // But we passed `isFirstRow` based on `index === 0`.
    // If `index === 0` it used style 1.
    // So we need to detect "First Data Row".
    // Actually `generatorRows` passes `index`. `isFirstRow` is `index === 0`.

    const isFirstDataRow = (rowIndex === 2); // 1-based rowIndex 2 means 2nd row (index 1)
    const isLastDataRow = isLastRow;

    if (isFirstDataRow && isFirstCol) styleIndex = 2;
    else if (isFirstDataRow && isLastCol) styleIndex = 4;
    else if (isFirstDataRow) styleIndex = 3;
    else if (isLastDataRow && isFirstCol) styleIndex = 8;
    else if (isLastDataRow && isLastCol) styleIndex = 10;
    else if (isLastDataRow) styleIndex = 9;
    else if (isFirstCol) styleIndex = 5;
    else if (isLastCol) styleIndex = 7;
    else styleIndex = 0; // Center / Inner
  }

  const style = ` s="${styleIndex}"`;

  return (
    cell.type === CELL_TYPE_STRING
      ? generatorStringCell(index, cell, rowIndex, style)
      : generatorNumberCell(index, cell, rowIndex, style)
  );
};
