package com.xzy.util;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import static org.junit.Assert.*;

public class XlsxSheetReaderTest {
    private static final String NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    @Rule
    public TemporaryFolder temporary = new TemporaryFolder();

    @Test
    public void headersAndDataKeepActualColumnsIncludingEmptyCells() throws Exception {
        File file = workbook("<row r=\"1\">" + inline("A1", "ID") + inline("C1", "Name")
                + "<c r=\"E1\"/></row>"
                + "<row r=\"2\">" + inline("A2", "one") + inline("C2", "three")
                + "<c r=\"E2\"><v>1.0</v></c></row>", false);

        assertEquals(Arrays.asList("ID", "", "Name", "", ""), XlsxSheetReader.readHeaders(file));
        assertEquals(Arrays.asList("one", "", "three", "", "1.0"), rows(file).get(0));
        assertEquals(XlsxSheetReader.readHeaders(file), ExcelUtils.readHeadersOnly(file));
    }

    @Test
    public void onlyCachedFormulaValuesAndInlineTextBecomeCellValues() throws Exception {
        File file = workbook("<row r=\"1\">" + inline("A1", "Header") + "</row>"
                + "<row r=\"2\">"
                + "<c r=\"A2\"><f>1+2</f><v>3.0</v></c>"
                + "<c r=\"B2\" t=\"str\"><f>CONCAT(&quot;a&quot;,&quot;b&quot;)</f><v>ab</v></c>"
                + "<c r=\"C2\" t=\"str\"><f>&quot;&quot;</f><v/></c>"
                + "<c r=\"D2\" t=\"b\"><f>1=1</f><v>1</v></c>"
                + "<c r=\"E2\" t=\"e\"><f>1/0</f><v>#DIV/0!</v></c>"
                + "<c r=\"F2\"><v>1.0</v></c>"
                + "<c r=\"G2\" t=\"inlineStr\"><is>\n<r><t> Foo </t></r>\n<r><t>Bar </t></r>"
                + "<rPh sb=\"0\" eb=\"1\"><t>phonetic</t></rPh>\n</is></c>"
                + "<c r=\"H2\" t=\"s\"><v>0</v></c>"
                + "<c r=\"I2\" t=\"b\"><f>1=0</f><v>0</v></c>"
                + "<c r=\"J2\"><f>25</f><v>2.5E1</v></c>"
                + "<c r=\"K2\"><v>2.5E1</v></c>"
                + "<c r=\"L2\" t=\"b\"><v>1</v></c><c r=\"M2\" t=\"b\"><v>0</v></c></row>", false);

        assertEquals(Arrays.asList("3", "ab", "", "true", "#DIV/0!", "1.0", "Foo Bar",
                "  shared value  ", "false", "25", "2.5E1", "1", "0"), rows(file).get(0));
    }

    @Test
    public void missingFormulaCachesReportTheirLocationButDoNotBlockHeaderLoading() throws Exception {
        for (String cache : Arrays.asList("", "<v/>", "<v> </v>")) {
            File file = workbook("<row r=\"1\">" + inline("A1", "Header") + "</row>"
                    + "<row r=\"7\"><c r=\"C7\"><f>1+2</f>" + cache + "</c></row>", false);

            assertEquals(Arrays.asList("Header"), XlsxSheetReader.readHeaders(file));
            IOException error = assertThrows(IOException.class, () -> rows(file));
            assertTrue(error.getMessage(), error.getMessage().contains(file.getAbsolutePath()));
            assertTrue(error.getMessage(), error.getMessage().contains("Data"));
            assertTrue(error.getMessage(), error.getMessage().contains("C7"));
            assertTrue(error.getMessage(), error.getMessage().contains("重新计算并保存"));
        }
    }

    @Test
    public void stringFormulaRequiresAnExplicitCacheEvenWhenItsResultCouldBeEmpty() throws Exception {
        File file = workbook("<row r=\"1\">" + inline("A1", "Header") + "</row>"
                + "<row r=\"2\"><c r=\"A2\" t=\"str\"><f>&quot;&quot;</f></c></row>", false);

        IOException error = assertThrows(IOException.class, () -> rows(file));
        assertTrue(error.getMessage(), error.getMessage().contains("A2"));
        assertTrue(error.getMessage(), error.getMessage().contains("公式缺少"));
    }

    @Test
    public void formulaInHeaderAlsoRequiresACachedResult() throws Exception {
        File file = workbook("<row r=\"1\"><c r=\"B1\" t=\"str\"><f>&quot;Header&quot;</f></c></row>", false);

        IOException error = assertThrows(IOException.class, () -> XlsxSheetReader.readHeaders(file));
        assertTrue(error.getMessage(), error.getMessage().contains("B1"));
    }

    @Test
    public void sharedFormulasUseEachCellsOwnCache() throws Exception {
        File file = workbook("<row r=\"1\">" + inline("A1", "Header") + "</row>"
                + "<row r=\"2\"><c r=\"A2\"><f t=\"shared\" ref=\"A2:A3\" si=\"0\">ROW()</f><v>2</v></c></row>"
                + "<row r=\"3\"><c r=\"A3\"><f t=\"shared\" si=\"0\"/><v>3</v></c></row>", false);

        assertEquals(Arrays.asList(Arrays.asList("2"), Arrays.asList("3")), rows(file));
    }

    @Test
    public void sharedStringsDoNotIncludePhoneticAnnotations() throws Exception {
        String xml = "<worksheet xmlns=\"" + NAMESPACE + "\"><sheetData>"
                + "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>"
                + "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>0</v></c></row>"
                + "</sheetData></worksheet>";
        String strings = "<sst xmlns=\"" + NAMESPACE + "\" count=\"2\" uniqueCount=\"1\"><si>"
                + "<r><t>汉</t></r><r><t>字</t></r><rPh sb=\"0\" eb=\"2\"><t>hànzì</t></rPh></si></sst>";
        File file = workbookXml(xml, false, strings);

        assertEquals(Arrays.asList("汉字"), XlsxSheetReader.readHeaders(file));
        assertEquals(Arrays.asList("汉字"), rows(file).get(0));
    }

    @Test
    public void datesRespectWorkbookDateSystemForOrdinaryAndFormulaCells() throws Exception {
        String data = "<row r=\"1\">" + inline("A1", "Header") + "</row>"
                + "<row r=\"2\"><c r=\"A2\" s=\"1\"><v>1</v></c>"
                + "<c r=\"B2\" s=\"1\"><f>1</f><v>1</v></c></row>";

        assertEquals(Arrays.asList("1900/01/01", "1900/01/01"), rows(workbook(data, false)).get(0));
        assertEquals(Arrays.asList("1904/01/02", "1904/01/02"), rows(workbook(data, true)).get(0));
    }

    @Test
    public void namespacePrefixesAndImplicitConsecutiveCellsAreSupported() throws Exception {
        String xml = "<x:worksheet xmlns:x=\"" + NAMESPACE + "\"><x:sheetData>"
                + "<x:row r=\"1\"><x:c r=\"C1\" t=\"inlineStr\"><x:is><x:t>Header</x:t></x:is></x:c></x:row>"
                + "<x:row r=\"2\"><x:c r=\"C2\"><x:v>3</x:v></x:c><x:c><x:v>4</x:v></x:c></x:row>"
                + "</x:sheetData></x:worksheet>";
        File file = workbookXml(xml, false);

        assertEquals(Arrays.asList("", "", "Header"), XlsxSheetReader.readHeaders(file));
        assertEquals(Arrays.asList("", "", "3", "4"), rows(file).get(0));
    }

    @Test
    public void sheetXmlCannotDeclareADoctype() throws Exception {
        String xml = "<!DOCTYPE worksheet>"
                + "<worksheet xmlns=\"" + NAMESPACE + "\"><sheetData><row r=\"1\">"
                + inline("A1", "Header") + "</row></sheetData></worksheet>";
        File file = workbookXml(xml, false);

        IOException error = assertThrows(IOException.class, () -> XlsxSheetReader.readHeaders(file));
        assertTrue(error.getMessage(), error.getMessage().contains("DOCTYPE"));
    }

    private File workbook(String sheetData, boolean date1904) throws IOException {
        return workbookXml("<worksheet xmlns=\"" + NAMESPACE + "\"><sheetData>"
                + sheetData + "</sheetData></worksheet>", date1904);
    }

    private File workbookXml(String sheetXml, boolean date1904) throws IOException {
        return workbookXml(sheetXml, date1904, null);
    }

    private File workbookXml(String sheetXml, boolean date1904, String sharedStringsXml) throws IOException {
        ByteArrayOutputStream original = new ByteArrayOutputStream();
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Data").createRow(0).createCell(0).setCellValue("  shared value  ");
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat((short) 14);
            workbook.getCTWorkbook().getWorkbookPr().setDate1904(date1904);
            workbook.write(original);
        }
        File file = temporary.newFile("sheet-" + System.nanoTime() + ".xlsx");
        try (ZipInputStream input = new ZipInputStream(new ByteArrayInputStream(original.toByteArray()));
             ZipOutputStream output = new ZipOutputStream(new FileOutputStream(file))) {
            ZipEntry entry;
            byte[] buffer = new byte[8192];
            while ((entry = input.getNextEntry()) != null) {
                output.putNextEntry(new ZipEntry(entry.getName()));
                if ("xl/worksheets/sheet1.xml".equals(entry.getName())) {
                    output.write(sheetXml.getBytes(StandardCharsets.UTF_8));
                } else if ("xl/sharedStrings.xml".equals(entry.getName()) && sharedStringsXml != null) {
                    output.write(sharedStringsXml.getBytes(StandardCharsets.UTF_8));
                } else {
                    int length;
                    while ((length = input.read(buffer)) >= 0) {
                        output.write(buffer, 0, length);
                    }
                }
                output.closeEntry();
            }
        }
        return file;
    }

    private static String inline(String reference, String text) {
        return "<c r=\"" + reference + "\" t=\"inlineStr\"><is><t>" + text + "</t></is></c>";
    }

    private static List<List<String>> rows(File file) throws IOException {
        List<List<String>> rows = new ArrayList<>();
        XlsxSheetReader.readDataRows(file, rows::add);
        return rows;
    }
}
