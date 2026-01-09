
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

export default (config) => {
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
