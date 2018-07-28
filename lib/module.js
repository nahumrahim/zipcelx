import escape from 'lodash.escape';
import JSZip from 'jszip';
import FileSaver from 'file-saver';

const CELL_TYPE_STRING = 'string';
const CELL_TYPE_NUMBER = 'number';
const validTypes = [CELL_TYPE_STRING, CELL_TYPE_NUMBER];

const MISSING_KEY_FILENAME = 'Zipclex config missing property filename';
const INVALID_TYPE_FILENAME = 'Zipclex filename can only be of type string';
const INVALID_TYPE_SHEET = 'Zipcelx sheet data is not of type array';
const INVALID_TYPE_SHEET_DATA = 'Zipclex sheet data childs is not of type array';

const WARNING_INVALID_TYPE = 'Invalid type supplied in cell config, falling back to "string"';

const childValidator = (array) => {
  return array.every(item => Array.isArray(item));
};

var validator = (config) => {
  if (!config.filename) {
    console.error(MISSING_KEY_FILENAME);
    return false;
  }

  if (typeof config.filename !== 'string') {
    console.error(INVALID_TYPE_FILENAME);
    return false;
  }

  if (!Array.isArray(config.sheet.data)) {
    console.error(INVALID_TYPE_SHEET);
    return false;
  }

  if (!childValidator(config.sheet.data)) {
    console.error(INVALID_TYPE_SHEET_DATA);
    return false;
  }

  return true;
};

const generateColumnLetter = (colIndex) => {
  if (typeof colIndex !== 'number') {
    return '';
  }

  const prefix = Math.floor(colIndex / 26);
  const letter = String.fromCharCode(97 + (colIndex % 26)).toUpperCase();
  if (prefix === 0) {
    return letter;
  }
  return generateColumnLetter(prefix - 1) + letter;
};

var generatorCellNumber = (index, rowNumber) => (
  `${generateColumnLetter(index)}${rowNumber}`
);

var generatorStringCell = (index, cell, rowIndex) => {
    var cStyle = '';
    if (cell.isHeader)
        cStyle = 's="1"';
    else if (cell.isFooter)
        cStyle = 's="2"';
    else if (cell.isSimple)
        cStyle = 's="5"';

    return (`<c r="${generatorCellNumber(index, rowIndex)}" t="inlineStr" ${cStyle}><is><t>${escape(cell.value)}</t></is></c>`)
};

var generatorNumberCell = (index, cell, rowIndex) => {
    var cStyle = '';
    if (cell.isHeader)
        cStyle = 's="1"';
    else if (cell.isFooter) {
        cStyle = 's="2"';
        if (cell.isMoney)
            cStyle = 's="4"';
    }
    else if (cell.isMoney)
        cStyle = 's="3"';
    else if (cell.isSimple)
        cStyle = 's="5"';
    
    return (`<c r="${generatorCellNumber(index, rowIndex)}" ${cStyle}><v>${cell.value}</v></c>`)
};

var formatCell = (cell, index, rowIndex) => {
  if (validTypes.indexOf(cell.type) === -1) {
    console.warn(WARNING_INVALID_TYPE);
    cell.type = CELL_TYPE_STRING;
  }

  return (
    cell.type === CELL_TYPE_STRING
    ? generatorStringCell(index, cell, rowIndex)
    : generatorNumberCell(index, cell, rowIndex)
  );
};

var formatRow = (row, index) => {
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

var generatorRows = rows => (
  rows
  .map((row, index) => formatRow(row, index))
  .join('')
);

var workbookXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
	<workbookPr/>
	<sheets>
		<sheet state="visible" name="Sheet1" sheetId="1" r:id="rId3"/>
	</sheets>
	<definedNames/>
	<calcPr/>
</workbook>`;

var workbookXMLRels = `<?xml version="1.0" ?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"/>
<Relationship Id="rId3" Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
</Relationships>`;

var rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

var contentTypes = `<?xml version="1.0" ?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default ContentType="application/xml" Extension="xml"/>
<Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/>
<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet1.xml"/>
<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" PartName="/xl/workbook.xml"/>
<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" PartName="/xl/styles.xml"/>
</Types>`;

var templateSheet = `<?xml version="1.0" ?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
<cols>{placeHolderCols}</cols>
<sheetData>{placeholder}</sheetData></worksheet>`;

// cellXfs
// s="0" default with wrap
// s="1" header
// s="2" footer
// s="4" footer-money
// s="3" money
// s="5" Simple: no format, no wrap
var templateStyles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" 
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
    <numFmts count="1">
        <numFmt numFmtId="165" formatCode="_-[$Q-100A]* #,##0.00_-;\-[$Q-100A]* #,##0.00_-;_-[$Q-100A]* &quot;-&quot;??_-;_-@_-"/>
    </numFmts>
    <fonts count="1" x14ac:knownFonts="1">
        <font>
            <sz val="12"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
        <font>
            <sz val="12"/>
            <color theme="0"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
        <font>
            <b/>
            <sz val="12"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
    </fonts>
    <fills count="4">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FF426ab3"/><bgColor indexed="64"/></patternFill></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFABB2B9"/><bgColor indexed="64"/></patternFill></fill>
    </fills>
    <borders count="3">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
        <border>
            <left/>
            <right/>
            <top style="thin">
                <color indexed="64"/>
            </top>
            <bottom/>
            <diagonal/>
        </border>
        <border>
            <left style="thin">
                <color indexed="64"/>
            </left>
            <right style="thin">
                <color indexed="64"/>
            </right>
            <top style="thin">
                <color indexed="64"/>
            </top>
            <bottom style="thin">
                <color indexed="64"/>
            </bottom>
            <diagonal/>
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="6">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1">
          <alignment vertical="center" wrapText="1" />
        </xf>
        <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1">
          <alignment vertical="center" wrapText="1"/>
        </xf>
        <xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1">
            <alignment vertical="center" wrapText="1" />
        </xf>
        <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1">
            <alignment vertical="center" wrapText="1"/>
        </xf>
        <xf numFmtId="165" fontId="2" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1">
            <alignment vertical="center" wrapText="1"/>
        </xf>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
            <alignment vertical="center" />
        </xf>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
    <extLst>
        <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" 
            xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
        </ext>
    </extLst>
</styleSheet>`;

const generateXMLWorksheet = (config, rows) => {
  const XMLRows = generatorRows(rows);
  var XMLCols = '';
  if (rows && rows.length > 0) {
    var XMLCol = '';
    for (var col = 1; col <= rows[0].length; col++) {
      XMLCol = '<col min="' + col + '" max="' + col + '" width="25" style="3" customWidth="1"/>';
      if (rows[0][col-1].width) {
        XMLCol = XMLCol.replace('width="25"', 'width="' + rows[0][col-1].width + '"');
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

var zipcelx = (config) => {
  if (!validator(config)) {
    throw new Error('Validation failed.');
  }

  const zip = new JSZip();
  const xl = zip.folder('xl');
  xl.file('workbook.xml', workbookXML);
  xl.file('_rels/workbook.xml.rels', workbookXMLRels);
  zip.file('_rels/.rels', rels);
  zip.file('[Content_Types].xml', contentTypes);

  var stylesSheet = templateStyles;
  if (config.headerBackground) {
    stylesSheet = stylesSheet.replace("FF426ab3", "FF" + config.headerBackground);
  }
  xl.file('styles.xml', stylesSheet);

  const worksheet = generateXMLWorksheet(config, config.sheet.data);
  xl.file('worksheets/sheet1.xml', worksheet);

  return zip.generateAsync({ type: 'blob' })
    .then((blob) => {
      FileSaver.saveAs(blob, `${config.filename}.xlsx`);
    });
};

export default zipcelx;
export { generateXMLWorksheet };
