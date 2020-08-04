# Development Notes

**Goal**: have a templating engine for Microsoft Word files. Avoid using existing
Rust libraries for DOCX manipulation (both poiscript and bokuweb) - instead do
non-destructive XML manipulation based on a few ground rules.

## Complex Templates

**Goal**: Be able to render complex data (like tables) into Word files using
Handlebars syntax. For example, the placeholder `{{table myTableData}}` should
render into a DOCX table.

**Specifics**: Normal templates inherit the paragraph attributes their are
placed in. For example, if the placeholder `{{placeholder}}` is placed in
paragraph, it would be represented in `word/document.xml` of the DOCX data as:

```xml
<w:p w14:paraId="186FFB15" w14:textId="76FEB211" w:rsidR="00CC5863" w:rsidRDefault="00CC5863">
    <w:r>
        <w:t>{{placeholder}}</w:t>
    </w:r>
</w:p>
```

If we pass two paragraphs of text (separated by two new line characters) to be
rendered using this template, the results would be the following:

```xml
<w:p w14:paraId="186FFB15" w14:textId="76FEB211" w:rsidR="00CC5863" w:rsidRDefault="00CC5863">
    <w:r>
        <w:t>Paragraph 1</w:t>
    </w:r>
</w:p>
<w:p w14:paraId="186FFB15" w14:textId="76FEB211" w:rsidR="00CC5863" w:rsidRDefault="00CC5863">
    <w:r>
        <w:t>Paragraph 2</w:t>
    </w:r>
</w:p>
```

So normal template placeholders retain the surrounding paragraph markup when rendered.

However, this might not be the case for more complex scenarios. For example, if we are rendering
data as a table, the output should be.

```xml
<w:tbl>
    <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
    </w:tblPr>
    <w:tblGrid>
        <w:gridCol w:w="4675"/>
        <w:gridCol w:w="4675"/>
    </w:tblGrid>
    <w:tr w:rsidR="00AD4011" w14:paraId="5A897CED" w14:textId="77777777" w:rsidTr="00AD4011">
        <w:tc>
            <w:tcPr>
                <w:tcW w:w="4675" w:type="dxa"/>
            </w:tcPr>
            <w:p w14:paraId="4F22606A" w14:textId="11B61EEC" w:rsidR="00AD4011" w:rsidRPr="00AD4011" w:rsidRDefault="00AD4011">
                <w:pPr>
                    <w:rPr>
                        <w:b/>
                        <w:bCs/>
                    </w:rPr>
                </w:pPr>
                <w:r w:rsidRPr="00AD4011">
                    <w:rPr>
                        <w:b/>
                        <w:bCs/>
                    </w:rPr>
                    <w:t>One</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:tcW w:w="4675" w:type="dxa"/>
            </w:tcPr>
            <w:p w14:paraId="74D9972D" w14:textId="54AACEB7" w:rsidR="00AD4011" w:rsidRPr="00AD4011" w:rsidRDefault="00AD4011">
                <w:pPr>
                    <w:rPr>
                        <w:b/>
                        <w:bCs/>
                    </w:rPr>
                </w:pPr>
                <w:r w:rsidRPr="00AD4011">
                    <w:rPr>
                        <w:b/>
                        <w:bCs/>
                    </w:rPr>
                    <w:t>Two</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>
    <w:tr w:rsidR="00AD4011" w14:paraId="48711FCA" w14:textId="77777777" w:rsidTr="00AD4011">
        <w:tc>
            <w:tcPr>
                <w:tcW w:w="4675" w:type="dxa"/>
            </w:tcPr>
            <w:p w14:paraId="678A795F" w14:textId="41A1E14D" w:rsidR="00AD4011" w:rsidRDefault="00AD4011">
                <w:r>
                    <w:t>Three</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:tcW w:w="4675" w:type="dxa"/>
            </w:tcPr>
            <w:p w14:paraId="465F1338" w14:textId="00E45351" w:rsidR="00AD4011" w:rsidRDefault="00AD4011">
                <w:r>
                    <w:t>Four</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>
</w:tbl>
``
```
