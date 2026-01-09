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

var generatorStringCell = (index, cell, rowIndex, style) => {
    return (`<c r="${generatorCellNumber(index, rowIndex)}" t="inlineStr"${style}><is><t>${escape(cell.value)}</t></is></c>`)
};

var generatorNumberCell = (index, cell, rowIndex, style) => {
    return (`<c r="${generatorCellNumber(index, rowIndex)}"${style}><v>${cell.value}</v></c>`)
};

var formatCell = (cell, index, rowIndex, isFirstRow, isLastRow, isFirstCol, isLastCol) => {
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

var formatRow = (row, index, totalRows, rowHeight, headerRowHeight) => {
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

var generatorRows = (rows, rowHeight, headerRowHeight) => (
  rows
    .map((row, index) => formatRow(row, index, rows.length, rowHeight, headerRowHeight))
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

const DEFAULT_FONT_SIZE = 12;
const DEFAULT_FONT_NAME = 'Calibri';

const generateFills = (headerFill, dataFill) => {
    let fills = [
        '<fill><patternFill patternType="none"/></fill>',
        '<fill><patternFill patternType="gray125"/></fill>'
    ];

    // Fill 2: Header Fill (Custom or Default)
    if (headerFill && headerFill.patternType === 'solid' && headerFill.fgColor) {
        fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="${headerFill.fgColor.replace('#', 'FF')}"/><bgColor indexed="64"/></patternFill></fill>`);
    } else {
        fills.push('<fill><patternFill patternType="solid"><fgColor rgb="FF426ab3"/><bgColor indexed="64"/></patternFill></fill>');
    }

    return `<fills count="${fills.length}">${fills.join('')}</fills>`;
};

const getBorderXml = (border, type) => {
    if (!border) return `<${type}/>`;
    return `<${type} style="${border.style || 'thin'}"><color rgb="${(border.color || 'FF000000').replace('#', 'FF')}"/></${type}>`;
};

const generateBorders = (headerBorder = {}, dataOutline = {}) => {
    // Border 0: None
    const borderNone = '<border><left/><right/><top/><bottom/><diagonal/></border>';

    // Border 1: Header (Custom or Default)
    let headerB = '';
    if (Object.keys(headerBorder).length > 0) {
        headerB = `<border>
            ${getBorderXml(headerBorder.left, 'left')}
            ${getBorderXml(headerBorder.right, 'right')}
            ${getBorderXml(headerBorder.top, 'top')}
            ${getBorderXml(headerBorder.bottom, 'bottom')}
            <diagonal/>
        </border>`;
    } else {
        // Default header border logic matching original if no config provided? 
        // Original didn't really have specific header borders separate from defaults often, relying on default styles. 
        // Let's assume standard thin black if not specified or just empty if we want to mimic "no border" default?
        // Actually the legacy code had `borderId="1"` for header. Let's look at legacy `styles.xml.js`.
        // Legacy `borderId="1"` was Top Thin. `borderId="2"` was All Thin.

        // For this new system, we will explicitly define:
        // If no header border config, we can default to ALL sides thin (like legacy style 2) or whatever user prefers. 
        // User asked for "apply all 4 styling for the borders".
        headerB = '<border><left/><right/><top/><bottom/><diagonal/></border>';
    }

    // Data Outline Borders (9 combinations)
    // We need to define borders for: 
    // 2: Top-Left Corner
    // 3: Top Edge
    // 4: Top-Right Corner
    // 5: Left Edge
    // 6: Center (No border or grid if requested?) -> User said "single square that contains all data", implying inner cells have NO border.
    // 7: Right Edge
    // 8: Bottom-Left Corner
    // 9: Bottom Edge
    // 10: Bottom-Right Corner

    // Helper for data outline
    const outline = dataOutline;

    const getB = (top, bottom, left, right) => {
        return `<border>
         ${getBorderXml(left ? outline : null, 'left')}
         ${getBorderXml(right ? outline : null, 'right')}
         ${getBorderXml(top ? outline : null, 'top')}
         ${getBorderXml(bottom ? outline : null, 'bottom')}
         <diagonal/>
      </border>`;
    };

    const borders = [
        borderNone, // 0
        headerB,    // 1
        getB(true, false, true, false), // 2: Top-Left
        getB(true, false, false, false), // 3: Top
        getB(true, false, false, true), // 4: Top-Right
        getB(false, false, true, false), // 5: Left
        getB(false, false, false, false), // 6: Center (inner) - effectively None
        getB(false, false, false, true), // 7: Right
        getB(false, true, true, false), // 8: Bottom-Left
        getB(false, true, false, false), // 9: Bottom
        getB(false, true, false, true)  // 10: Bottom-Right
    ];

    return `<borders count="${borders.length}">${borders.join('')}</borders>`;
};

const generateCellXfs = () => {
    // Style XFs mapping
    // We map our abstract "Style IDs" to these XFs.

    // 0: Default (No style) - Uses Border 0, Fill 0
    // 1: Header - Uses Border 1, Fill 2 (Header Fill)
    // 2-10: Data Styles - Use Borders 2-10, Fill 0 (None)

    let xfs = [];

    // 0: Default
    xfs.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment vertical="center" wrapText="1" /></xf>');

    // 1: Header
    xfs.push('<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFill="1" applyBorder="1"><alignment vertical="center" wrapText="1"/></xf>');

    // 2-10: Data Edge Styles
    // We need 9 styles for the data block matrix passed to the cell formatter
    for (let i = 2; i <= 10; i++) {
        xfs.push(`<xf numFmtId="0" fontId="0" fillId="0" borderId="${i}" xfId="0" applyBorder="1"><alignment vertical="center" wrapText="1"/></xf>`);
    }

    return `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>`;
};

var styleGenerator = (config) => {
    const styleConfig = config.style || {};
    const headerFill = styleConfig.header ? styleConfig.header.fill : null;
    const headerBorder = styleConfig.header ? styleConfig.header.border : {};
    const dataOutline = styleConfig.data ? styleConfig.data.outline : {};

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" 
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
    <numFmts count="1">
        <numFmt numFmtId="165" formatCode="_-[$Q-100A]* #,##0.00_-;\\-[$Q-100A]* #,##0.00_-;_-[$Q-100A]* &quot;-&quot;??_-;_-@_-"/>
    </numFmts>
    <fonts count="2" x14ac:knownFonts="1">
        <font>
            <sz val="${DEFAULT_FONT_SIZE}"/>
            <color theme="1"/>
            <name val="${DEFAULT_FONT_NAME}"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
        <font>
            <b/>
            <sz val="${DEFAULT_FONT_SIZE}"/>
            <color theme="1"/>
            <name val="${DEFAULT_FONT_NAME}"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
    </fonts>
    ${generateFills(headerFill)}
    ${generateBorders(headerBorder, dataOutline)}
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    ${generateCellXfs()}
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
};

const generateXMLWorksheet = (config, rows) => {
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

export default zipcelx;
export { generateXMLWorksheet };
