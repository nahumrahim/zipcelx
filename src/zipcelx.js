import JSZip from 'jszip';
import FileSaver from 'file-saver';

import validator from './validator';
import generatorRows from './formatters/rows/generatorRows';

import workbookXML from './statics/workbook.xml';
import workbookXMLRels from './statics/workbook.xml.rels';
import rels from './statics/rels';
import contentTypes from './statics/[Content_Types].xml';
import templateSheet from './templates/worksheet.xml';
import styleGenerator from './commons/styleGenerator';

export const generateXMLWorksheet = (config, rows) => {
  const rowHeight = config.style && config.style.row ? config.style.row.height : undefined;
  const headerRowHeight = config.style && config.style.header ? config.style.header.height : undefined;
  const XMLRows = generatorRows(rows, rowHeight, headerRowHeight);
  var XMLCols = '';
  if (rows && rows.length > 0) {
    var XMLCol = '';
    for (var col = 1; col <= rows[0].length; col++) {
      XMLCol = '<col min="' + col + '" max="' + col + '" width="25" style="0" customWidth="1"/>';
      if (rows[0][col - 1].width) {
        XMLCol = XMLCol.replace('width="25"', 'width="' + rows[0][col - 1].width + '"');
      }
      XMLCols += XMLCol;
    }
  }
  const worksheetOutput =
    templateSheet
      .replace('{placeHolderCols}', XMLCols)
      .replace('{placeholder}', XMLRows);

  return worksheetOutput;
};

export default (config) => {
  if (!validator(config)) {
    throw new Error('Validation failed.');
  }

  const zip = new JSZip();
  const xl = zip.folder('xl');
  xl.file('workbook.xml', workbookXML);
  xl.file('_rels/workbook.xml.rels', workbookXMLRels);
  zip.file('_rels/.rels', rels);
  zip.file('[Content_Types].xml', contentTypes);

  const stylesSheet = styleGenerator(config);
  xl.file('styles.xml', stylesSheet);

  const worksheet = generateXMLWorksheet(config, config.sheet.data);
  xl.file('worksheets/sheet1.xml', worksheet);

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  })
    .then((blob) => {
      FileSaver.saveAs(blob, `${config.filename}.xlsx`);
    });
};

