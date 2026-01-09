import escape from 'lodash.escape';
import generatorCellNumber from '../../commons/generatorCellNumber';

export default (index, cell, rowIndex, style) => {
    return (`<c r="${generatorCellNumber(index, rowIndex)}" t="inlineStr"${style}><is><t>${escape(cell.value)}</t></is></c>`)
};
