import generatorCellNumber from '../../commons/generatorCellNumber';

export default (index, cell, rowIndex) => {
    var cStyle = ''
    if (cell.isHeader)
        cStyle = 's="1"'
    else if (cell.isFooter) {
        cStyle = 's="2"'
        if (cell.isMoney)
            cStyle = 's="4"'
    }
    else if (cell.isMoney)
        cStyle = 's="3"'
    else if (cell.isSimple)
        cStyle = 's="5"'
    
    return (`<c r="${generatorCellNumber(index, rowIndex)}" ${cStyle}><v>${cell.value}</v></c>`)
};