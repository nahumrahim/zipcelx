import formatRow from '../../../src/formatters/rows/formatRow';
import baseConfig from '../../baseConfig';

const expectedXML = '<row r="1"><c r="A1" t="inlineStr" s="1"><is><t>Test</t></is></c><c r="B1" s="1"><v>1000</v></c></row>';

describe('Format Row', () => {
  it('Should create one row from given data', () => {
    expect(formatRow(baseConfig.sheet.data[0], 0, 1)).toBe(expectedXML);
  });

  it('Should apply custom row height', () => {
    const customHeightXML = '<row ht="30" customHeight="1" r="1"><c r="A1" t="inlineStr" s="1"><is><t>Test</t></is></c><c r="B1" s="1"><v>1000</v></c></row>';
    expect(formatRow(baseConfig.sheet.data[0], 0, 1, 30)).toBe(customHeightXML);
  });

  it('Should apply custom header height', () => {
    // Header row is index 0. Default is 25.
    const headerRow = [
      { value: 'Test', type: 'string', isHeader: true },
      { value: 1000, type: 'number', isHeader: true }
    ];
    // Note: s="1" is applied because it's the first row (index 0) OR because isHeader is true.
    // In formatCell.js: if (cell.isHeader || isFirstRow) styleIndex = 1;
    const customHeaderHeightXML = '<row ht="40" customHeight="1" r="1"><c r="A1" t="inlineStr" s="1"><is><t>Test</t></is></c><c r="B1" s="1"><v>1000</v></c></row>';
    expect(formatRow(headerRow, 0, 1, undefined, 40)).toBe(customHeaderHeightXML);
  });
});
