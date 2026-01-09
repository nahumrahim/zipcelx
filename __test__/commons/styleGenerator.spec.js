import styleGenerator from '../../src/commons/styleGenerator';

describe('Style Generator', () => {
    it('Should generate default styles when no config is provided', () => {
        const xml = styleGenerator({});
        expect(xml).toContain('<cellXfs count="11">'); // 0 + 1 + 9
        expect(xml).toContain('<fill><patternFill patternType="solid"><fgColor rgb="FF426ab3"/><bgColor indexed="64"/></patternFill></fill>'); // Default header fill
    });

    it('Should generate custom header fill', () => {
        const config = {
            style: {
                header: {
                    fill: { patternType: 'solid', fgColor: 'FFFF0000' }
                }
            }
        };
        const xml = styleGenerator(config);
        expect(xml).toContain('<fgColor rgb="FFFF0000"/>');
    });

    it('Should generate data outline borders', () => {
        const config = {
            style: {
                data: {
                    outline: { style: 'thin', color: 'FF00FF00' }
                }
            }
        };
        const xml = styleGenerator(config);
        // Check for one of the borders, e.g. Border 2 (Top-Left)
        // It should have left and top borders
        expect(xml).toContain('<left style="thin"><color rgb="FF00FF00"/></left>');
        expect(xml).toContain('<top style="thin"><color rgb="FF00FF00"/></top>');
    });
});
