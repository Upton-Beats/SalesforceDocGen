/**
 * DocGenExcelServiceTest
 *
 * Tests for DocGenExcelService using in-memory .xlsx (OOXML ZIP) templates.
 */
@IsTest
private class DocGenExcelServiceTest {

    // ─────────────────────────────────────────────────────────────────────────
    // Template helpers
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Builds a minimal valid .xlsx ZIP blob with merge tags in the worksheet.
     */
    private static Blob buildXlsxTemplate(String sheetRowsXml) {
        String contentTypes =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
            '</Types>';

        String rels =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
            '</Relationships>';

        String workbook =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
            '</workbook>';

        String workbookRels =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
            '</Relationships>';

        String sheet =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            '<sheetData>' +
            sheetRowsXml +
            '</sheetData>' +
            '</worksheet>';

        Compression.ZipWriter zw = new Compression.ZipWriter();
        zw.addEntry('[Content_Types].xml', Blob.valueOf(contentTypes));
        zw.addEntry('_rels/.rels', Blob.valueOf(rels));
        zw.addEntry('xl/workbook.xml', Blob.valueOf(workbook));
        zw.addEntry('xl/_rels/workbook.xml.rels', Blob.valueOf(workbookRels));
        zw.addEntry('xl/worksheets/sheet1.xml', Blob.valueOf(sheet));
        return zw.getArchive();
    }

    /**
     * Builds a .xlsx with shared strings table.
     */
    private static Blob buildXlsxWithSharedStrings(List<String> sharedStrings, String sheetRowsXml) {
        String contentTypes =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
            '</Types>';

        String rels =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
            '</Relationships>';

        String workbook =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
            '</workbook>';

        String workbookRels =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
            '</Relationships>';

        String sheet =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            '<sheetData>' +
            sheetRowsXml +
            '</sheetData>' +
            '</worksheet>';

        String ssXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' +
            sharedStrings.size() + '" uniqueCount="' + sharedStrings.size() + '">';
        for (String s : sharedStrings) {
            ssXml += '<si><t>' + s + '</t></si>';
        }
        ssXml += '</sst>';

        Compression.ZipWriter zw = new Compression.ZipWriter();
        zw.addEntry('[Content_Types].xml', Blob.valueOf(contentTypes));
        zw.addEntry('_rels/.rels', Blob.valueOf(rels));
        zw.addEntry('xl/workbook.xml', Blob.valueOf(workbook));
        zw.addEntry('xl/_rels/workbook.xml.rels', Blob.valueOf(workbookRels));
        zw.addEntry('xl/worksheets/sheet1.xml', Blob.valueOf(sheet));
        zw.addEntry('xl/sharedStrings.xml', Blob.valueOf(ssXml));
        return zw.getArchive();
    }

    /**
     * Extracts worksheet XML from a generated .xlsx blob.
     */
    private static String extractSheetXml(Blob xlsxBlob) {
        Compression.ZipReader reader = new Compression.ZipReader(xlsxBlob);
        for (Compression.ZipEntry entry : reader.getEntries()) {
            if (entry.getName() == 'xl/worksheets/sheet1.xml') {
                return reader.extract(entry.getName()).toString();
            }
        }
        return null;
    }

    /**
     * Extracts shared strings XML from a generated .xlsx blob.
     */
    private static String extractSharedStringsXml(Blob xlsxBlob) {
        Compression.ZipReader reader = new Compression.ZipReader(xlsxBlob);
        for (Compression.ZipEntry entry : reader.getEntries()) {
            if (entry.getName() == 'xl/sharedStrings.xml') {
                return reader.extract(entry.getName()).toString();
            }
        }
        return null;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Tests
    // ─────────────────────────────────────────────────────────────────────────

    @IsTest
    static void testSimpleTagReplacement() {
        String rows =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c>' +
            '<c r="B1" t="inlineStr"><is><t>{Name}</t></is></c></row>' +
            '<row r="2"><c r="A2" t="inlineStr"><is><t>Status</t></is></c>' +
            '<c r="B2" t="inlineStr"><is><t>{Status}</t></is></c></row>';

        Blob templateBlob = buildXlsxTemplate(rows);
        Map<String, Object> data = new Map<String, Object>{
            'Name' => 'Acme Corp',
            'Status' => 'Active'
        };

        Blob result = DocGenExcelService.processExcelTemplate(templateBlob, data);
        String sheetXml = extractSheetXml(result);

        System.assert(sheetXml.contains('Acme Corp'), 'Name should be replaced');
        System.assert(sheetXml.contains('Active'), 'Status should be replaced');
        System.assert(!sheetXml.contains('{Name}'), 'Tag should be removed');
        System.assert(!sheetXml.contains('{Status}'), 'Tag should be removed');
    }

    @IsTest
    static void testSharedStringReplacement() {
        List<String> sharedStrings = new List<String>{ 'Label', '{Name}', '{Industry}' };
        String rows =
            '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>' +
            '<row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2" t="s"><v>2</v></c></row>';

        Blob templateBlob = buildXlsxWithSharedStrings(sharedStrings, rows);
        Map<String, Object> data = new Map<String, Object>{
            'Name' => 'Test Co',
            'Industry' => 'Technology'
        };

        Blob result = DocGenExcelService.processExcelTemplate(templateBlob, data);
        String ssXml = extractSharedStringsXml(result);

        System.assert(ssXml.contains('Test Co'), 'Name should be in shared strings');
        System.assert(ssXml.contains('Technology'), 'Industry should be in shared strings');
    }

    @IsTest
    static void testXmlEscaping() {
        String rows =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>{CompanyName}</t></is></c></row>';

        Blob templateBlob = buildXlsxTemplate(rows);
        Map<String, Object> data = new Map<String, Object>{
            'CompanyName' => 'Smith & Jones <LLC>'
        };

        Blob result = DocGenExcelService.processExcelTemplate(templateBlob, data);
        String sheetXml = extractSheetXml(result);

        System.assert(sheetXml.contains('Smith &amp; Jones &lt;LLC&gt;'), 'Special characters should be XML-escaped');
    }

    @IsTest
    static void testNullValueRemovesTag() {
        String rows =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>{MissingField}</t></is></c></row>';

        Blob templateBlob = buildXlsxTemplate(rows);
        Map<String, Object> data = new Map<String, Object>{
            'MissingField' => null
        };

        Blob result = DocGenExcelService.processExcelTemplate(templateBlob, data);
        String sheetXml = extractSheetXml(result);

        System.assert(!sheetXml.contains('{MissingField}'), 'Null value tag should be removed');
    }

    @IsTest
    static void testDateFormatting() {
        String rows =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>{MyDate}</t></is></c></row>';

        Blob templateBlob = buildXlsxTemplate(rows);
        Map<String, Object> data = new Map<String, Object>{
            'MyDate' => Date.newInstance(2025, 3, 15)
        };

        Blob result = DocGenExcelService.processExcelTemplate(templateBlob, data);
        String sheetXml = extractSheetXml(result);

        System.assert(sheetXml.contains('03/15/2025'), 'Date should be formatted MM/dd/yyyy');
    }

    @IsTest
    static void testDateWithCustomFormat() {
        String rows =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>{MyDate:yyyy-MM-dd}</t></is></c></row>';

        Blob templateBlob = buildXlsxTemplate(rows);
        Map<String, Object> data = new Map<String, Object>{
            'MyDate' => Date.newInstance(2025, 7, 4)
        };

        Blob result = DocGenExcelService.processExcelTemplate(templateBlob, data);
        String sheetXml = extractSheetXml(result);

        System.assert(sheetXml.contains('2025-07-04'), 'Date should be formatted with custom pattern');
    }

    @IsTest
    static void testDottedFieldPath() {
        String rows =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>{Account.Name}</t></is></c></row>';

        Blob templateBlob = buildXlsxTemplate(rows);
        Map<String, Object> data = new Map<String, Object>{
            'Account' => new Map<String, Object>{ 'Name' => 'Parent Corp' }
        };

        Blob result = DocGenExcelService.processExcelTemplate(templateBlob, data);
        String sheetXml = extractSheetXml(result);

        System.assert(sheetXml.contains('Parent Corp'), 'Dotted path should resolve');
    }

    @IsTest
    static void testPreservesNonWorksheetEntries() {
        String rows =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>{Name}</t></is></c></row>';

        // Build template with an extra media file (simulating header/footer image)
        Compression.ZipWriter zw = new Compression.ZipWriter();

        String contentTypes =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Default Extension="png" ContentType="image/png"/>' +
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
            '</Types>';

        String sheet =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            '<sheetData>' + rows + '</sheetData></worksheet>';

        zw.addEntry('[Content_Types].xml', Blob.valueOf(contentTypes));
        zw.addEntry('xl/worksheets/sheet1.xml', Blob.valueOf(sheet));
        zw.addEntry('xl/media/image1.png', Blob.valueOf('FAKEPNGDATA'));
        Blob templateBlob = zw.getArchive();

        Map<String, Object> data = new Map<String, Object>{ 'Name' => 'Test' };

        Blob result = DocGenExcelService.processExcelTemplate(templateBlob, data);

        // Verify image entry survived
        Compression.ZipReader reader = new Compression.ZipReader(result);
        Boolean foundImage = false;
        for (Compression.ZipEntry entry : reader.getEntries()) {
            if (entry.getName() == 'xl/media/image1.png') {
                String content = reader.extract(entry.getName()).toString();
                System.assertEquals('FAKEPNGDATA', content, 'Image data should be preserved');
                foundImage = true;
            }
        }
        System.assert(foundImage, 'Image entry should exist in output');
    }

    @IsTest
    static void testNullTemplateThrowsException() {
        try {
            DocGenExcelService.processExcelTemplate(null, new Map<String, Object>());
            System.assert(false, 'Should have thrown exception');
        } catch (DocGenException e) {
            System.assert(e.getMessage().contains('missing'), 'Should mention missing template');
        }
    }

    @IsTest
    static void testEmptyDataReturnsCleanOutput() {
        String rows =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>{Name}</t></is></c></row>';

        Blob templateBlob = buildXlsxTemplate(rows);
        Map<String, Object> data = new Map<String, Object>();

        Blob result = DocGenExcelService.processExcelTemplate(templateBlob, data);
        String sheetXml = extractSheetXml(result);

        // Unmatched tag should be removed (resolves to null)
        System.assert(!sheetXml.contains('{Name}'), 'Unmatched tag should be removed');
    }

    @IsTest
    static void testRowLoopWithChildRecords() {
        String rows =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row>' +
            '<row r="2"><c r="A2" t="inlineStr"><is><t>{#Items}</t></is></c>' +
            '<c r="B2" t="inlineStr"><is><t>{ProductName}</t></is></c>' +
            '<c r="C2" t="inlineStr"><is><t>{Quantity}</t></is></c>' +
            '<c r="D2" t="inlineStr"><is><t>{/Items}</t></is></c></row>' +
            '<row r="3"><c r="A3" t="inlineStr"><is><t>Footer</t></is></c></row>';

        Blob templateBlob = buildXlsxTemplate(rows);
        Map<String, Object> data = new Map<String, Object>{
            'Items' => new Map<String, Object>{
                'records' => new List<Object>{
                    new Map<String, Object>{ 'ProductName' => 'Widget A', 'Quantity' => 10 },
                    new Map<String, Object>{ 'ProductName' => 'Widget B', 'Quantity' => 5 }
                }
            }
        };

        Blob result = DocGenExcelService.processExcelTemplate(templateBlob, data);
        String sheetXml = extractSheetXml(result);

        System.assert(sheetXml.contains('Widget A'), 'First child record should appear');
        System.assert(sheetXml.contains('Widget B'), 'Second child record should appear');
        System.assert(sheetXml.contains('10'), 'Quantity should appear');
        System.assert(sheetXml.contains('5'), 'Quantity should appear');
        System.assert(sheetXml.contains('Header'), 'Header row should be preserved');
        System.assert(sheetXml.contains('Footer'), 'Footer row should be preserved');
    }

    @IsTest
    static void testRowLoopWithListData() {
        String rows =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>{#LineItems}</t></is></c>' +
            '<c r="B1" t="inlineStr"><is><t>{Name}</t></is></c>' +
            '<c r="C1" t="inlineStr"><is><t>{/LineItems}</t></is></c></row>';

        Blob templateBlob = buildXlsxTemplate(rows);
        Map<String, Object> data = new Map<String, Object>{
            'LineItems' => new List<Object>{
                new Map<String, Object>{ 'Name' => 'Item 1' },
                new Map<String, Object>{ 'Name' => 'Item 2' },
                new Map<String, Object>{ 'Name' => 'Item 3' }
            }
        };

        Blob result = DocGenExcelService.processExcelTemplate(templateBlob, data);
        String sheetXml = extractSheetXml(result);

        System.assert(sheetXml.contains('Item 1'), 'First item should appear');
        System.assert(sheetXml.contains('Item 2'), 'Second item should appear');
        System.assert(sheetXml.contains('Item 3'), 'Third item should appear');
    }

    @IsTest
    static void testCaseInsensitiveFieldResolution() {
        String rows =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>{name}</t></is></c></row>';

        Blob templateBlob = buildXlsxTemplate(rows);
        Map<String, Object> data = new Map<String, Object>{
            'Name' => 'Case Test'
        };

        Blob result = DocGenExcelService.processExcelTemplate(templateBlob, data);
        String sheetXml = extractSheetXml(result);

        System.assert(sheetXml.contains('Case Test'), 'Case-insensitive resolution should work');
    }

    @IsTest
    static void testEmptyLoopRemovesRow() {
        String rows =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row>' +
            '<row r="2"><c r="A2" t="inlineStr"><is><t>{#EmptyList}</t></is></c>' +
            '<c r="B2" t="inlineStr"><is><t>{Name}</t></is></c>' +
            '<c r="C2" t="inlineStr"><is><t>{/EmptyList}</t></is></c></row>';

        Blob templateBlob = buildXlsxTemplate(rows);
        Map<String, Object> data = new Map<String, Object>{
            'EmptyList' => new List<Object>()
        };

        Blob result = DocGenExcelService.processExcelTemplate(templateBlob, data);
        String sheetXml = extractSheetXml(result);

        System.assert(sheetXml.contains('Header'), 'Header should be preserved');
        // The loop row should be omitted since no records
        System.assert(!sheetXml.contains('{#EmptyList}'), 'Loop tags should be removed');
    }
}
