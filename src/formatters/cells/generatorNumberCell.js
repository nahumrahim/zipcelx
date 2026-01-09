import generatorCellNumber from '../../commons/generatorCellNumber';

export default (index, cell, rowIndex, style) => {
    return (`<c r="${generatorCellNumber(index, rowIndex)}"${style}><v>${cell.value}</v></c>`)
};
