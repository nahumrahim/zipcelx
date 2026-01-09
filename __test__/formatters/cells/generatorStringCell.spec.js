import generatorStringCell from '../../../src/formatters/cells/generatorStringCell';

const expectedXML = '<c r="A1" t="inlineStr" s="0"><is><t>Test</t></is></c>';

describe('Cell of type String', () => {
  it('Should create a new xml markup cell', () => {
    expect(generatorStringCell(0, { value: 'Test' }, 1, ' s="0"')).toBe(expectedXML);
  });
});
