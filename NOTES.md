# Development Notes

**Goal**: have a templating engine for Microsoft Word files. Avoid using existing
Rust libraries for DOCX manipulation (both poiscript and bokuweb) - instead do
non-destructive XML manipulation based on a few ground rules.

_N.B!_ A very good (and relatively short)
[introduction](https://www.toptal.com/xml/an-informal-introduction-to-docx) to
the DOCX format by Stepan Yakovenko.

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

## Bugs With Broken Up Text Runs

```
<w:p w14:paraId="0EF41D1A" w14:textId="4124C809" w:rsidR="00A56FB5" w:rsidRDefault="00A56FB5">
    <w:r>
        <w:t>{{</w:t>
    </w:r>
    <w:proofErr w:type="spellStart"/>
    <w:r>
        <w:t>myPlaceholder</w:t>
    </w:r>
    <w:proofErr w:type="spellEnd"/>
    <w:r>
        <w:t>}}</w:t>
    </w:r>
</w:p>
```

**Fixed in v0.1.3.**

## Jupyter Notebooks

**Goal**: Add the ability to render Jupyter Notebooks down into Word files by
doing something like `{{jupyter my_notebook}}`.

Good news: Jupyter Notebooks are just long JSON files.

## Adding Images

Here's how to add images to a docx file, AFAIK:

1. Add an entry to the `word/_rels/document.xml.rels` file containing the name
   of the image file. Like so:
   ```xml
   <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        ...
        <Relationship
            Id="rId4"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="media/image1.png"/>
   </Relationships>
   ```
2. Add the file as `media/image1.png` to the DOCX zip.
3. Add the following markup to the `word/document.xml`:
   ```xml
      <w:p w14:paraId="6BAD9101" w14:textId="0DE1D2E6" w:rsidR="00F6259F" w:rsidRDefault="00121631">
         <w:r>
            <w:rPr>
               <w:noProof/>
            </w:rPr>
            <w:drawing>
               <wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="2B64807F" wp14:editId="3291176C">
                  <wp:extent cx="5943600" cy="4243705"/>
                  <wp:effectExtent l="0" t="0" r="0" b="4445"/>
                  <wp:docPr id="1" name="Picture 1" descr="A picture containing bird, tree, flower&#xA;&#xA;Description automatically generated"/>
                  <wp:cNvGraphicFramePr>
                        <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                  </wp:cNvGraphicFramePr>
                  <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                           <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                              <pic:nvPicPr>
                                    <pic:cNvPr id="1" name="Picture 1" descr="A picture containing bird, tree, flower&#xA;&#xA;Description automatically generated"/>
                                    <pic:cNvPicPr/>
                              </pic:nvPicPr>
                              <pic:blipFill>
                                    <a:blip r:embed="rId4">
                                       <a:extLst>
                                          <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                                                <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                                          </a:ext>
                                       </a:extLst>
                                    </a:blip>
                                    <a:stretch>
                                       <a:fillRect/>
                                    </a:stretch>
                              </pic:blipFill>
                              <pic:spPr>
                                    <a:xfrm>
                                       <a:off x="0" y="0"/>
                                       <a:ext cx="5943600" cy="4243705"/>
                                    </a:xfrm>
                                    <a:prstGeom prst="rect">
                                       <a:avLst/>
                                    </a:prstGeom>
                              </pic:spPr>
                           </pic:pic>
                        </a:graphicData>
                  </a:graphic>
               </wp:inline>
            </w:drawing>
      </w:r>
   </w:p>
   ```
