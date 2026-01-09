import formatRow from './formatRow';

export default (rows, rowHeight, headerRowHeight) => (
  rows
    .map((row, index) => formatRow(row, index, rows.length, rowHeight, headerRowHeight))
    .join('')
);
