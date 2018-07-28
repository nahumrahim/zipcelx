// cellXfs
// s="0" default with wrap
// s="1" header
// s="2" footer
// s="4" footer-money
// s="3" money
// s="5" Simple: no format, no wrap
export default `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
