import generatorNumberCell from '../../../src/formatters/cells/generatorNumberCell';

export const expectedXML = '<c r="A1" s="0"><v>1000</v></c>';

describe('Cell of type Number', () => {
  it('Should create a new xml markup cell', () => {
    expect(generatorNumberCell(0, { value: 1000 }, 1, ' s="0"')).toBe(expectedXML);
  });
});
