// Helper functions to generate Office Open XML for Excel charts

export function generateLineChartXML(
    sheetName: string,
    dataStartRow: number,
    dataEndRow: number,
    categoryColumn: string, // e.g., 'A'
    valueColumn: string,    // e.g., 'D'
    chartTitle: string,
    seriesName: string,
    lineColor: string       // hex without #
): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" 
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <c:date1904 val="0"/>
    <c:lang val="en-US"/>
    <c:roundedCorners val="0"/>
    <c:chart>
        <c:title>
            <c:tx>
                <c:rich>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:pPr><a:defRPr/></a:pPr>
                        <a:r>
                            <a:rPr lang="en-US" b="1" sz="1400"/>
                            <a:t>${chartTitle}</a:t>
                        </a:r>
                    </a:p>
                </c:rich>
            </c:tx>
            <c:layout/>
        </c:title>
        <c:autoTitleDeleted val="0"/>
        <c:plotArea>
            <c:layout/>
            <c:lineChart>
                <c:grouping val="standard"/>
                <c:ser>
                    <c:idx val="0"/>
                    <c:order val="0"/>
                    <c:tx>
                        <c:v>${seriesName}</c:v>
                    </c:tx>
                    <c:spPr>
                        <a:ln w="28575">
                            <a:solidFill>
                                <a:srgbClr val="${lineColor}"/>
                            </a:solidFill>
                        </a:ln>
                    </c:spPr>
                    <c:marker>
                        <c:symbol val="circle"/>
                        <c:size val="5"/>
                    </c:marker>
                    <c:cat>
                        <c:strRef>
                            <c:f>'${sheetName}'!$${categoryColumn}$${dataStartRow}:$${categoryColumn}$${dataEndRow}</c:f>
                        </c:strRef>
                    </c:cat>
                    <c:val>
                        <c:numRef>
                            <c:f>'${sheetName}'!$${valueColumn}$${dataStartRow}:$${valueColumn}$${dataEndRow}</c:f>
                        </c:numRef>
                    </c:val>
                    <c:smooth val="0"/>
                </c:ser>
                <c:dLbls>
                    <c:showLegendKey val="0"/>
                    <c:showVal val="0"/>
                    <c:showCatName val="0"/>
                    <c:showSerName val="0"/>
                    <c:showPercent val="0"/>
                    <c:showBubbleSize val="0"/>
                </c:dLbls>
                <c:axId val="84580096"/>
                <c:axId val="84582144"/>
            </c:lineChart>
            <c:catAx>
                <c:axId val="84580096"/>
                <c:scaling>
                    <c:orientation val="minMax"/>
                </c:scaling>
                <c:delete val="0"/>
                <c:axPos val="b"/>
                <c:title>
                    <c:tx>
                        <c:rich>
                            <a:bodyPr/>
                            <a:lstStyle/>
                            <a:p>
                                <a:pPr><a:defRPr/></a:pPr>
                                <a:r>
                                    <a:rPr lang="en-US"/>
                                    <a:t>Month-Year</a:t>
                                </a:r>
                            </a:p>
                        </c:rich>
                    </c:tx>
                    <c:layout/>
                </c:title>
                <c:numFmt formatCode="General" sourceLinked="1"/>
                <c:majorTickMark val="out"/>
                <c:minorTickMark val="none"/>
                <c:tickLblPos val="nextTo"/>
                <c:crossAx val="84582144"/>
                <c:crosses val="autoZero"/>
                <c:auto val="1"/>
                <c:lblAlgn val="ctr"/>
                <c:lblOffset val="100"/>
            </c:catAx>
            <c:valAx>
                <c:axId val="84582144"/>
                <c:scaling>
                    <c:orientation val="minMax"/>
                </c:scaling>
                <c:delete val="0"/>
                <c:axPos val="l"/>
                <c:majorGridlines/>
                <c:title>
                    <c:tx>
                        <c:rich>
                            <a:bodyPr/>
                            <a:lstStyle/>
                            <a:p>
                                <a:pPr><a:defRPr/></a:pPr>
                                <a:r>
                                    <a:rPr lang="en-US"/>
                                    <a:t>Amount ($)</a:t>
                                </a:r>
                            </a:p>
                        </c:rich>
                    </c:tx>
                    <c:layout/>
                </c:title>
                <c:numFmt formatCode="General" sourceLinked="1"/>
                <c:majorTickMark val="out"/>
                <c:minorTickMark val="none"/>
                <c:tickLblPos val="nextTo"/>
                <c:crossAx val="84580096"/>
                <c:crosses val="autoZero"/>
                <c:crossBetween val="between"/>
            </c:valAx>
        </c:plotArea>
        <c:legend>
            <c:legendPos val="r"/>
            <c:layout/>
        </c:legend>
        <c:plotVisOnly val="1"/>
        <c:dispBlanksAs val="gap"/>
        <c:showDLblsOverMax val="0"/>
    </c:chart>
    <c:printSettings>
        <c:headerFooter/>
        <c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>
        <c:pageSetup/>
    </c:printSettings>
</c:chartSpace>`;
}

export function generateDrawingXML(chartId: string = "rId1"): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" 
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <xdr:twoCellAnchor editAs="oneCell">
        <xdr:from>
            <xdr:col>0</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>25</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
        </xdr:from>
        <xdr:to>
            <xdr:col>7</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>40</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
        </xdr:to>
        <xdr:graphicFrame macro="">
            <xdr:nvGraphicFramePr>
                <xdr:cNvPr id="2" name="Chart 1"/>
                <xdr:cNvGraphicFramePr/>
            </xdr:nvGraphicFramePr>
            <xdr:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="0" cy="0"/>
            </xdr:xfrm>
            <a:graphic>
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                    <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" 
                             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                             r:id="${chartId}"/>
                </a:graphicData>
            </a:graphic>
        </xdr:graphicFrame>
        <xdr:clientData/>
    </xdr:twoCellAnchor>
</xdr:wsDr>`;
}

export function generateDrawingRelsXML(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`;
}

export function generateComposedChartXML(
    sheetName: string,
    dataStartRow: number,
    dataEndRow: number,
    categoryColumn: string,
    barColumn: string,      // Price Change %
    lineColumn: string,     // Sales Change %
    chartTitle: string,
    barSeriesName: string,
    lineSeriesName: string,
    barColor: string,       // hex without #
    lineColor: string       // hex without #
): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" 
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <c:date1904 val="0"/>
    <c:lang val="en-US"/>
    <c:roundedCorners val="0"/>
    <c:chart>
        <c:title>
            <c:tx>
                <c:rich>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:pPr><a:defRPr/></a:pPr>
                        <a:r>
                            <a:rPr lang="en-US" b="1" sz="1400"/>
                            <a:t>${chartTitle}</a:t>
                        </a:r>
                    </a:p>
                </c:rich>
            </c:tx>
            <c:layout/>
        </c:title>
        <c:autoTitleDeleted val="0"/>
        <c:plotArea>
            <c:layout/>
            <c:barChart>
                <c:barDir val="col"/>
                <c:grouping val="clustered"/>
                <c:ser>
                    <c:idx val="0"/>
                    <c:order val="0"/>
                    <c:tx>
                        <c:v>${barSeriesName}</c:v>
                    </c:tx>
                    <c:spPr>
                        <a:solidFill>
                            <a:srgbClr val="${barColor}"/>
                        </a:solidFill>
                    </c:spPr>
                    <c:cat>
                        <c:strRef>
                            <c:f>'${sheetName}'!$${categoryColumn}$${dataStartRow}:$${categoryColumn}$${dataEndRow}</c:f>
                        </c:strRef>
                    </c:cat>
                    <c:val>
                        <c:numRef>
                            <c:f>'${sheetName}'!$${barColumn}$${dataStartRow}:$${barColumn}$${dataEndRow}</c:f>
                        </c:numRef>
                    </c:val>
                </c:ser>
                <c:dLbls>
                    <c:showLegendKey val="0"/>
                    <c:showVal val="0"/>
                    <c:showCatName val="0"/>
                    <c:showSerName val="0"/>
                    <c:showPercent val="0"/>
                    <c:showBubbleSize val="0"/>
                </c:dLbls>
                <c:axId val="84580096"/>
                <c:axId val="84582144"/>
            </c:barChart>
            <c:lineChart>
                <c:grouping val="standard"/>
                <c:ser>
                    <c:idx val="1"/>
                    <c:order val="1"/>
                    <c:tx>
                        <c:v>${lineSeriesName}</c:v>
                    </c:tx>
                    <c:spPr>
                        <a:ln w="28575">
                            <a:solidFill>
                                <a:srgbClr val="${lineColor}"/>
                            </a:solidFill>
                        </a:ln>
                    </c:spPr>
                    <c:marker>
                        <c:symbol val="circle"/>
                        <c:size val="5"/>
                        <c:spPr>
                            <a:solidFill>
                                <a:srgbClr val="${lineColor}"/>
                            </a:solidFill>
                        </c:spPr>
                    </c:marker>
                    <c:cat>
                        <c:strRef>
                            <c:f>'${sheetName}'!$${categoryColumn}$${dataStartRow}:$${categoryColumn}$${dataEndRow}</c:f>
                        </c:strRef>
                    </c:cat>
                    <c:val>
                        <c:numRef>
                            <c:f>'${sheetName}'!$${lineColumn}$${dataStartRow}:$${lineColumn}$${dataEndRow}</c:f>
                        </c:numRef>
                    </c:val>
                </c:ser>
                <c:dLbls>
                    <c:showLegendKey val="0"/>
                    <c:showVal val="0"/>
                    <c:showCatName val="0"/>
                    <c:showSerName val="0"/>
                    <c:showPercent val="0"/>
                    <c:showBubbleSize val="0"/>
                </c:dLbls>
                <c:axId val="84580096"/>
                <c:axId val="84582144"/>
            </c:lineChart>
            <c:catAx>
                <c:axId val="84580096"/>
                <c:scaling>
                    <c:orientation val="minMax"/>
                </c:scaling>
                <c:delete val="0"/>
                <c:axPos val="b"/>
                <c:title>
                    <c:tx>
                        <c:rich>
                            <a:bodyPr/>
                            <a:lstStyle/>
                            <a:p>
                                <a:pPr><a:defRPr/></a:pPr>
                                <a:r>
                                    <a:rPr lang="en-US"/>
                                    <a:t>Month-Year</a:t>
                                </a:r>
                            </a:p>
                        </c:rich>
                    </c:tx>
                    <c:layout/>
                </c:title>
                <c:numFmt formatCode="General" sourceLinked="1"/>
                <c:majorTickMark val="out"/>
                <c:minorTickMark val="none"/>
                <c:tickLblPos val="nextTo"/>
                <c:crossAx val="84582144"/>
                <c:crosses val="autoZero"/>
                <c:auto val="1"/>
                <c:lblAlgn val="ctr"/>
                <c:lblOffset val="100"/>
            </c:catAx>
            <c:valAx>
                <c:axId val="84582144"/>
                <c:scaling>
                    <c:orientation val="minMax"/>
                </c:scaling>
                <c:delete val="0"/>
                <c:axPos val="l"/>
                <c:majorGridlines/>
                <c:title>
                    <c:tx>
                        <c:rich>
                            <a:bodyPr/>
                            <a:lstStyle/>
                            <a:p>
                                <a:pPr><a:defRPr/></a:pPr>
                                <a:r>
                                    <a:rPr lang="en-US"/>
                                    <a:t>Change %</a:t>
                                </a:r>
                            </a:p>
                        </c:rich>
                    </c:tx>
                    <c:layout/>
                </c:title>
                <c:numFmt formatCode="0%" sourceLinked="0"/>
                <c:majorTickMark val="out"/>
                <c:minorTickMark val="none"/>
                <c:tickLblPos val="nextTo"/>
                <c:crossAx val="84580096"/>
                <c:crosses val="autoZero"/>
                <c:crossBetween val="between"/>
            </c:valAx>
        </c:plotArea>
        <c:legend>
            <c:legendPos val="r"/>
            <c:layout/>
        </c:legend>
        <c:plotVisOnly val="1"/>
        <c:dispBlanksAs val="gap"/>
        <c:showDLblsOverMax val="0"/>
    </c:chart>
    <c:printSettings>
        <c:headerFooter/>
        <c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>
        <c:pageSetup/>
    </c:printSettings>
</c:chartSpace>`;
}

export function generateMultiLineChartXML(
    sheetName: string,
    dataStartRow: number,
    dataEndRow: number,
    categoryColumn: string,
    series: Array<{
        name: string;
        dataColumn: string;
        color: string;
    }>,
    chartTitle: string,
    yAxisTitle: string
): string {
    const seriesXML = series.map((s, index) => `
                <c:ser>
                    <c:idx val="${index}"/>
                    <c:order val="${index}"/>
                    <c:tx>
                        <c:v>${s.name}</c:v>
                    </c:tx>
                    <c:spPr>
                        <a:ln w="28575">
                            <a:solidFill>
                                <a:srgbClr val="${s.color}"/>
                            </a:solidFill>
                        </a:ln>
                    </c:spPr>
                    <c:marker>
                        <c:symbol val="circle"/>
                        <c:size val="5"/>
                        <c:spPr>
                            <a:solidFill>
                                <a:srgbClr val="${s.color}"/>
                            </a:solidFill>
                        </c:spPr>
                    </c:marker>
                    <c:cat>
                        <c:strRef>
                            <c:f>'${sheetName}'!$${categoryColumn}$${dataStartRow}:$${categoryColumn}$${dataEndRow}</c:f>
                        </c:strRef>
                    </c:cat>
                    <c:val>
                        <c:numRef>
                            <c:f>'${sheetName}'!$${s.dataColumn}$${dataStartRow}:$${s.dataColumn}$${dataEndRow}</c:f>
                        </c:numRef>
                    </c:val>
                    <c:smooth val="0"/>
                </c:ser>`).join('');

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" 
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <c:date1904 val="0"/>
    <c:lang val="en-US"/>
    <c:roundedCorners val="0"/>
    <c:chart>
        <c:title>
            <c:tx>
                <c:rich>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:pPr><a:defRPr/></a:pPr>
                        <a:r>
                            <a:rPr lang="en-US" b="1" sz="1400"/>
                            <a:t>${chartTitle}</a:t>
                        </a:r>
                    </a:p>
                </c:rich>
            </c:tx>
            <c:layout/>
        </c:title>
        <c:autoTitleDeleted val="0"/>
        <c:plotArea>
            <c:layout/>
            <c:lineChart>
                <c:grouping val="standard"/>
                ${seriesXML}
                <c:dLbls>
                    <c:showLegendKey val="0"/>
                    <c:showVal val="0"/>
                    <c:showCatName val="0"/>
                    <c:showSerName val="0"/>
                    <c:showPercent val="0"/>
                    <c:showBubbleSize val="0"/>
                </c:dLbls>
                <c:axId val="84580096"/>
                <c:axId val="84582144"/>
            </c:lineChart>
            <c:catAx>
                <c:axId val="84580096"/>
                <c:scaling>
                    <c:orientation val="minMax"/>
                </c:scaling>
                <c:delete val="0"/>
                <c:axPos val="b"/>
                <c:title>
                    <c:tx>
                        <c:rich>
                            <a:bodyPr/>
                            <a:lstStyle/>
                            <a:p>
                                <a:pPr><a:defRPr/></a:pPr>
                                <a:r>
                                    <a:rPr lang="en-US"/>
                                    <a:t>Month-Year</a:t>
                                </a:r>
                            </a:p>
                        </c:rich>
                    </c:tx>
                    <c:layout/>
                </c:title>
                <c:numFmt formatCode="General" sourceLinked="1"/>
                <c:majorTickMark val="out"/>
                <c:minorTickMark val="none"/>
                <c:tickLblPos val="nextTo"/>
                <c:crossAx val="84582144"/>
                <c:crosses val="autoZero"/>
                <c:auto val="1"/>
                <c:lblAlgn val="ctr"/>
                <c:lblOffset val="100"/>
            </c:catAx>
            <c:valAx>
                <c:axId val="84582144"/>
                <c:scaling>
                    <c:orientation val="minMax"/>
                </c:scaling>
                <c:delete val="0"/>
                <c:axPos val="l"/>
                <c:majorGridlines/>
                <c:title>
                    <c:tx>
                        <c:rich>
                            <a:bodyPr/>
                            <a:lstStyle/>
                            <a:p>
                                <a:pPr><a:defRPr/></a:pPr>
                                <a:r>
                                    <a:rPr lang="en-US"/>
                                    <a:t>${yAxisTitle}</a:t>
                                </a:r>
                            </a:p>
                        </c:rich>
                    </c:tx>
                    <c:layout/>
                </c:title>
                <c:numFmt formatCode="0%" sourceLinked="0"/>
                <c:majorTickMark val="out"/>
                <c:minorTickMark val="none"/>
                <c:tickLblPos val="nextTo"/>
                <c:crossAx val="84580096"/>
                <c:crosses val="autoZero"/>
                <c:crossBetween val="between"/>
            </c:valAx>
        </c:plotArea>
        <c:legend>
            <c:legendPos val="r"/>
            <c:layout/>
        </c:legend>
        <c:plotVisOnly val="1"/>
        <c:dispBlanksAs val="gap"/>
        <c:showDLblsOverMax val="0"/>
    </c:chart>
    <c:printSettings>
        <c:headerFooter/>
        <c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>
        <c:pageSetup/>
    </c:printSettings>
</c:chartSpace>`;
}

export function generateMultiBarChartXML(
    sheetName: string,
    dataStartRow: number,
    dataEndRow: number,
    categoryColumn: string,
    series: Array<{
        name: string;
        dataColumn: string;
        color: string;
    }>,
    chartTitle: string,
    yAxisTitle: string
): string {
    const seriesXML = series.map((s, index) => `
                <c:ser>
                    <c:idx val="${index}"/>
                    <c:order val="${index}"/>
                    <c:tx>
                        <c:v>${s.name}</c:v>
                    </c:tx>
                    <c:spPr>
                        <a:solidFill>
                            <a:srgbClr val="${s.color}"/>
                        </a:solidFill>
                    </c:spPr>
                    <c:cat>
                        <c:strRef>
                            <c:f>'${sheetName}'!$${categoryColumn}$${dataStartRow}:$${categoryColumn}$${dataEndRow}</c:f>
                        </c:strRef>
                    </c:cat>
                    <c:val>
                        <c:numRef>
                            <c:f>'${sheetName}'!$${s.dataColumn}$${dataStartRow}:$${s.dataColumn}$${dataEndRow}</c:f>
                        </c:numRef>
                    </c:val>
                </c:ser>`).join('');

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" 
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <c:date1904 val="0"/>
    <c:lang val="en-US"/>
    <c:roundedCorners val="0"/>
    <c:chart>
        <c:title>
            <c:tx>
                <c:rich>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:pPr><a:defRPr/></a:pPr>
                        <a:r>
                            <a:rPr lang="en-US" b="1" sz="1400"/>
                            <a:t>${chartTitle}</a:t>
                        </a:r>
                    </a:p>
                </c:rich>
            </c:tx>
            <c:layout/>
        </c:title>
        <c:autoTitleDeleted val="0"/>
        <c:plotArea>
            <c:layout/>
            <c:barChart>
                <c:barDir val="col"/>
                <c:grouping val="clustered"/>
                ${seriesXML}
                <c:dLbls>
                    <c:showLegendKey val="0"/>
                    <c:showVal val="0"/>
                    <c:showCatName val="0"/>
                    <c:showSerName val="0"/>
                    <c:showPercent val="0"/>
                    <c:showBubbleSize val="0"/>
                </c:dLbls>
                <c:axId val="84580096"/>
                <c:axId val="84582144"/>
            </c:barChart>
            <c:catAx>
                <c:axId val="84580096"/>
                <c:scaling>
                    <c:orientation val="minMax"/>
                </c:scaling>
                <c:delete val="0"/>
                <c:axPos val="b"/>
                <c:title>
                    <c:tx>
                        <c:rich>
                            <a:bodyPr/>
                            <a:lstStyle/>
                            <a:p>
                                <a:pPr><a:defRPr/></a:pPr>
                                <a:r>
                                    <a:rPr lang="en-US"/>
                                    <a:t>Month-Year</a:t>
                                </a:r>
                            </a:p>
                        </c:rich>
                    </c:tx>
                    <c:layout/>
                </c:title>
                <c:numFmt formatCode="General" sourceLinked="1"/>
                <c:majorTickMark val="out"/>
                <c:minorTickMark val="none"/>
                <c:tickLblPos val="nextTo"/>
                <c:crossAx val="84582144"/>
                <c:crosses val="autoZero"/>
                <c:auto val="1"/>
                <c:lblAlgn val="ctr"/>
                <c:lblOffset val="100"/>
            </c:catAx>
            <c:valAx>
                <c:axId val="84582144"/>
                <c:scaling>
                    <c:orientation val="minMax"/>
                </c:scaling>
                <c:delete val="0"/>
                <c:axPos val="l"/>
                <c:majorGridlines/>
                <c:title>
                    <c:tx>
                        <c:rich>
                            <a:bodyPr/>
                            <a:lstStyle/>
                            <a:p>
                                <a:pPr><a:defRPr/></a:pPr>
                                <a:r>
                                    <a:rPr lang="en-US"/>
                                    <a:t>${yAxisTitle}</a:t>
                                </a:r>
                            </a:p>
                        </c:rich>
                    </c:tx>
                    <c:layout/>
                </c:title>
                <c:numFmt formatCode="0%" sourceLinked="0"/>
                <c:majorTickMark val="out"/>
                <c:minorTickMark val="none"/>
                <c:tickLblPos val="nextTo"/>
                <c:crossAx val="84580096"/>
                <c:crosses val="autoZero"/>
                <c:crossBetween val="between"/>
            </c:valAx>
        </c:plotArea>
        <c:legend>
            <c:legendPos val="r"/>
            <c:layout/>
        </c:legend>
        <c:plotVisOnly val="1"/>
        <c:dispBlanksAs val="gap"/>
        <c:showDLblsOverMax val="0"/>
    </c:chart>
    <c:printSettings>
        <c:headerFooter/>
        <c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>
        <c:pageSetup/>
    </c:printSettings>
</c:chartSpace>`;
}