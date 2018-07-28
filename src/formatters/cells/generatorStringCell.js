import escape from 'lodash.escape';
import generatorCellNumber from '../../commons/generatorCellNumber';

export default (index, cell, rowIndex) => {
    var cStyle = ''
    if (cell.isHeader)
        cStyle = 's="1"'
    else if (cell.isFooter)
        cStyle = 's="2"'
    else if (cell.isSimple)
        cStyle = 's="5"'

    return (`<c r="${generatorCellNumber(index, rowIndex)}" t="inlineStr" ${cStyle}><is><t>${escape(cell.value)}</t></is></c>`)
};