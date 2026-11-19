package com.xzy;

import com.xzy.util.ExcelUtils;
import com.xzy.util.FormulaCache;
import com.xzy.util.FormulaCache.FormulaCacheException;
import com.xzy.util.SafeExport;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Date1904Support;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Excel表格合并工具。
 */
public class TableMerger {
    private static final Logger logger = LogManager.getLogger(TableMerger.class);
    private static final int MEMORY_WINDOW_SIZE = 100;
    private static final String HEADER_DATA_SHEET_NAME = "合并数据";
    private static final String HEADER_REPORT_SHEET_NAME = "合并报告";
    private static final String WPS_DATA_SHEET_NAME = "总表";
    private static final String WPS_REPORT_SHEET_NAME = "报告";

    public enum MergeStrategy {
        AUTO,
        HEADER_BASED,
        WPS_STYLE_NO_HEADER
    }

    public static final class MergeSummary {
        private final int successCount;
        private final int partialCount;
        private final int failureCount;

        private MergeSummary(int successCount, int partialCount, int failureCount) {
            this.successCount = successCount;
            this.partialCount = partialCount;
            this.failureCount = failureCount;
        }

        public int getSuccessCount() { return successCount; }
        public int getPartialCount() { return partialCount; }
        public int getFailureCount() { return failureCount; }

        public String formatSummary() {
            return "成功 " + successCount + " / 部分成功 " + partialCount + " / 失败 " + failureCount;
        }
    }

    public static final class AmbiguousHeadersException extends IOException {
        private AmbiguousHeadersException() {
            super("无法可靠判断各表是否有表头。请选择“有表头合并”或“无表头合并”。无表头合并会保留首行数据。");
        }
    }

    /**
     * 合并单个Excel文件中的所有sheet页。
     */
    public void mergeSheetsInWorkbook(String inputFilePath, String outputFilePath,
                                      ProgressListener progressListener) throws IOException {
        mergeSheetsInWorkbook(inputFilePath, outputFilePath, MergeStrategy.AUTO, progressListener);
    }

    public void mergeSheetsInWorkbook(String inputFilePath, String outputFilePath,
                                      MergeStrategy strategy,
                                      ProgressListener progressListener) throws IOException {
        mergeSheetsInWorkbookWithSummary(inputFilePath, outputFilePath, strategy, progressListener);
    }

    public MergeSummary mergeSheetsInWorkbookWithSummary(String inputFilePath, String outputFilePath,
                                                         MergeStrategy strategy, ProgressListener progressListener)
            throws IOException {
        File inputFile = new File(inputFilePath);
        validateReadableFile(inputFile, "输入文件");
        File outputFile = prepareOutputFile(outputFilePath);

        try (SafeExport export = SafeExport.begin(outputFile, inputFile)) {
            MergeSummary summary;
            try (FileInputStream fis = new FileInputStream(inputFile);
                 Workbook inputWorkbook = WorkbookFactory.create(fis);
                 SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(MEMORY_WINDOW_SIZE)) {
                try {
                    MergeStrategy effectiveStrategy = strategy == null || strategy == MergeStrategy.AUTO
                            ? detectStrategy(collectSheetSamples(inputWorkbook)) : strategy;
                    logger.info("表格合并策略: {}", effectiveStrategy);
                    if (effectiveStrategy == MergeStrategy.WPS_STYLE_NO_HEADER) {
                        summary = mergeWorkbookSheetsWpsStyle(inputFile, inputWorkbook, outputWorkbook, progressListener);
                    } else {
                        summary = mergeWorkbookSheetsHeaderBased(inputFile, inputWorkbook, outputWorkbook, progressListener);
                    }
                    try (FileOutputStream fos = new FileOutputStream(export.temporaryFile())) {
                        outputWorkbook.write(fos);
                    }
                } finally {
                    outputWorkbook.dispose();
                }
            }
            export.commit();
            return summary;
        } catch (FormulaCacheException e) {
            throw e.inFile(inputFile);
        } catch (IOException e) {
            throw e;
        } catch (Exception e) {
            logger.error("合并工作簿失败", e);
            throw new IOException("合并工作簿失败: " + e.getMessage(), e);
        }
    }

    /**
     * 合并多个Excel文件。每个文件默认读取第一个sheet页。
     */
    public void mergeFiles(List<File> inputFiles, String outputFilePath,
                           ProgressListener progressListener) throws IOException {
        mergeFiles(inputFiles, outputFilePath, MergeStrategy.AUTO, progressListener);
    }

    public void mergeFiles(List<File> inputFiles, String outputFilePath,
                           MergeStrategy strategy,
                           ProgressListener progressListener) throws IOException {
        mergeFilesWithSummary(inputFiles, outputFilePath, strategy, progressListener);
    }

    public MergeSummary mergeFilesWithSummary(List<File> inputFiles, String outputFilePath,
                                              MergeStrategy strategy, ProgressListener progressListener) throws IOException {
        if (inputFiles == null || inputFiles.isEmpty()) {
            throw new IOException("请选择要合并的Excel文件");
        }
        File outputFile = prepareOutputFile(outputFilePath);

        try (SafeExport export = SafeExport.begin(outputFile, inputFiles.toArray(new File[0]))) {
            MergeStrategy effectiveStrategy = strategy == null || strategy == MergeStrategy.AUTO
                    ? detectStrategy(collectFileSamples(inputFiles)) : strategy;
            logger.info("表格合并策略: {}", effectiveStrategy);
            MergeSummary summary;
            if (effectiveStrategy == MergeStrategy.WPS_STYLE_NO_HEADER) {
                summary = mergeFilesWpsStyle(inputFiles, export.temporaryFile(), progressListener);
            } else {
                summary = mergeFilesHeaderBased(inputFiles, export.temporaryFile(), progressListener);
            }
            export.commit();
            return summary;
        }
    }

    private MergeSummary mergeWorkbookSheetsHeaderBased(File inputFile, Workbook inputWorkbook, SXSSFWorkbook outputWorkbook,
                                                        ProgressListener progressListener) throws IOException {
        HeaderMergeWriter writer = new HeaderMergeWriter(outputWorkbook);
        int total = inputWorkbook.getNumberOfSheets();
        for (int i = 0; i < total; i++) {
            processHeaderSheet(inputFile.getAbsolutePath(), inputWorkbook.getSheetAt(i), writer);
            notifyProgress(progressListener, i + 1, total);
        }
        return writer.finish();
    }

    private MergeSummary mergeFilesHeaderBased(List<File> inputFiles, File outputFile,
                                              ProgressListener progressListener) throws IOException {
        try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(MEMORY_WINDOW_SIZE)) {
            try {
                HeaderMergeWriter writer = new HeaderMergeWriter(outputWorkbook);
                int total = inputFiles.size();
                for (int i = 0; i < total; i++) {
                    processFirstSheetHeaderBased(inputFiles.get(i), writer);
                    notifyProgress(progressListener, i + 1, total);
                }
                MergeSummary summary = writer.finish();
                try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                    outputWorkbook.write(fos);
                }
                return summary;
            } finally {
                outputWorkbook.dispose();
            }
        } catch (IOException e) {
            throw e;
        } catch (Exception e) {
            logger.error("合并文件失败", e);
            throw new IOException("合并文件失败: " + e.getMessage(), e);
        }
    }

    private void processFirstSheetHeaderBased(File inputFile, HeaderMergeWriter writer) throws IOException {
        if (inputFile == null || !inputFile.exists() || !inputFile.canRead()) {
            String path = inputFile == null ? "" : inputFile.getAbsolutePath();
            writer.addReport(path, "", "失败", 0, 0, "无法读取文件");
            return;
        }

        Workbook source;
        try {
            source = WorkbookFactory.create(inputFile, null, true);
        } catch (Exception e) {
            logger.warn("读取文件失败: {}", inputFile.getAbsolutePath(), e);
            writer.addReport(inputFile.getAbsolutePath(), "", "失败", 0, 0, e.getMessage());
            return;
        }
        try (Workbook workbook = source) {
            if (workbook.getNumberOfSheets() == 0) {
                writer.addReport(inputFile.getAbsolutePath(), "", "失败", 0, 0, "未找到有效表头");
                return;
            }
            processHeaderSheet(inputFile.getAbsolutePath(), workbook.getSheetAt(0), writer);
        } catch (FormulaCacheException e) {
            throw e.inFile(inputFile);
        }
    }

    private void processHeaderSheet(String sourceFile, Sheet sourceSheet, HeaderMergeWriter writer)
            throws IOException {
        String sheetName = sourceSheet == null ? "" : sourceSheet.getSheetName();
        if (sourceSheet == null) {
            writer.addReport(sourceFile, sheetName, "失败", 0, 0, "未找到有效表头");
            return;
        }

        Iterator<Row> rows = sourceSheet.iterator();
        List<String> headers = readHeaders(firstHeaderRow(rows));
        if (headers.isEmpty()) {
            writer.addReport(sourceFile, sheetName, "失败", 0, 0, "未找到有效表头");
            return;
        }

        writer.ensureStandardHeaders(headers);
        SourceMapping mapping = buildSourceMapping(headers, writer.getStandardHeaders(), sourceFile + " / " + sheetName);
        if (mapping.targetToSourceIndex.isEmpty()) {
            writer.addReport(sourceFile, sheetName, "失败", 0, 0, "与标准表头没有共同列");
            return;
        }
        mapping.targetToSourceIndex.forEach((target, source) ->
                widenFromSource(sourceSheet, source, writer.dataSheet, target));
        int readRows = 0;
        int mergedRows = 0;
        while (rows.hasNext()) {
            Row row = rows.next();
            if (!hasAnyCellValue(row)) {
                continue;
            }
            readRows++;
            writer.appendDataRow(row, mapping.targetToSourceIndex);
            mergedRows++;
        }

        String status = mapping.missingHeaders.isEmpty() ? "成功" : "部分成功";
        String remark = mapping.missingHeaders.isEmpty() ? "" : "缺失列: " + String.join("、", mapping.missingHeaders);
        writer.addReport(sourceFile, sheetName, status, readRows, mergedRows, remark);
    }

    private MergeSummary mergeWorkbookSheetsWpsStyle(File inputFile, Workbook inputWorkbook, SXSSFWorkbook outputWorkbook,
                                                     ProgressListener progressListener) throws IOException {
        WpsMergeWriter writer = new WpsMergeWriter(outputWorkbook);
        int total = inputWorkbook.getNumberOfSheets();
        for (int i = 0; i < total; i++) {
            writer.processSheet(inputFile, 1, inputWorkbook.getSheetAt(i));
            notifyProgress(progressListener, i + 1, total);
        }
        return writer.finish();
    }

    private MergeSummary mergeFilesWpsStyle(List<File> inputFiles, File outputFile,
                                           ProgressListener progressListener) throws IOException {
        try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(MEMORY_WINDOW_SIZE)) {
            try {
                WpsMergeWriter writer = new WpsMergeWriter(outputWorkbook);
                int total = inputFiles.size();
                for (int i = 0; i < total; i++) {
                    processFirstSheetWps(inputFiles.get(i), i + 1, writer);
                    notifyProgress(progressListener, i + 1, total);
                }
                MergeSummary summary = writer.finish();
                try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                    outputWorkbook.write(fos);
                }
                return summary;
            } finally {
                outputWorkbook.dispose();
            }
        }
    }

    private void processFirstSheetWps(File inputFile, int workbookIndex, WpsMergeWriter writer) throws IOException {
        Workbook source;
        try {
            validateReadableFile(inputFile, "输入文件");
            source = WorkbookFactory.create(inputFile, null, true);
        } catch (Exception e) {
            logger.warn("读取文件失败: {}", inputFile, e);
            writer.addFailure(inputFile, workbookIndex, "", "失败：" + e.getMessage());
            return;
        }
        try (Workbook workbook = source) {
            if (workbook.getNumberOfSheets() == 0) {
                writer.addFailure(inputFile, workbookIndex, "", "失败：没有数据");
            } else {
                writer.processSheet(inputFile, workbookIndex, workbook.getSheetAt(0));
            }
        } catch (FormulaCacheException e) {
            throw e.inFile(inputFile);
        } catch (RuntimeException e) {
            throw new IOException("合并文件失败：" + inputFile.getAbsolutePath() + "：" + e.getMessage(), e);
        }
    }

    private List<SheetSample> collectSheetSamples(Workbook workbook) throws FormulaCacheException {
        List<SheetSample> samples = new ArrayList<>();
        for (Sheet sheet : workbook) {
            collectSample(sheet, samples);
        }
        return samples;
    }

    private List<SheetSample> collectFileSamples(List<File> inputFiles) throws IOException {
        List<SheetSample> samples = new ArrayList<>();
        for (File inputFile : inputFiles) {
            if (inputFile == null || !inputFile.exists() || !inputFile.canRead()) {
                continue;
            }
            try (FileInputStream fis = new FileInputStream(inputFile);
                 Workbook workbook = WorkbookFactory.create(fis)) {
                if (workbook.getNumberOfSheets() > 0) {
                    collectSample(workbook.getSheetAt(0), samples);
                }
            } catch (FormulaCacheException e) {
                throw e.inFile(inputFile);
            } catch (Exception e) {
                logger.warn("采集合并样本失败: {}", inputFile.getAbsolutePath(), e);
            }
        }
        return samples;
    }

    private MergeStrategy detectStrategy(List<SheetSample> samples) throws AmbiguousHeadersException {
        if (samples.isEmpty()) {
            return MergeStrategy.HEADER_BASED;
        }
        if (samples.stream().allMatch(sample -> sample.startsWithData)) {
            return MergeStrategy.WPS_STYLE_NO_HEADER;
        }
        if (samples.size() == 1 || samples.stream().anyMatch(sample -> sample.startsWithData)
                || samples.stream().noneMatch(sample -> sample.typedDataBelow)) {
            throw new AmbiguousHeadersException();
        }
        Set<String> base = samples.get(0).firstValues;
        for (int i = 1; i < samples.size(); i++) {
            Set<String> current = samples.get(i).firstValues;
            int smallerSize = Math.min(base.size(), current.size());
            if (current.stream().filter(base::contains).count() < smallerSize * 0.6) {
                throw new AmbiguousHeadersException();
            }
        }
        return MergeStrategy.HEADER_BASED;
    }

    private static class SheetSample {
        private final Set<String> firstValues = new HashSet<>();
        private boolean startsWithData;
        private boolean typedDataBelow;
    }

    private boolean looksLikeDataValue(String value) {
        String text = value == null ? "" : value.trim();
        if (text.matches(".*\\d{6,}.*")) {
            return true;
        }
        if (text.matches(".*\\d{4}[/-]\\d{1,2}.*")) {
            return true;
        }
        return text.matches("\\d+(\\.\\d+)?");
    }

    private void collectSample(Sheet sheet, List<SheetSample> samples) throws FormulaCacheException {
        Iterator<Row> rows = sheet.iterator();
        Row first = firstHeaderRow(rows);
        if (first == null) return;
        SheetSample sample = new SheetSample();
        sample.startsWithData = hasTypedData(first);
        for (Cell cell : first) {
            String text = readMergeCellText(cell);
            String field = normalizeHeader(text);
            if (!field.isEmpty()) {
                sample.firstValues.add(field);
                if (looksLikeDataValue(text)) sample.startsWithData = true;
            }
        }
        while (rows.hasNext()) {
            Row row = rows.next();
            if (hasAnyCellValue(row)) {
                sample.typedDataBelow = hasTypedData(row);
                break;
            }
        }
        samples.add(sample);
    }

    private Row firstHeaderRow(Iterator<Row> rows) throws FormulaCacheException {
        while (rows.hasNext()) {
            Row row = rows.next();
            for (Cell cell : row) {
                if (!normalizeHeader(readMergeCellText(cell)).isEmpty()) return row;
            }
        }
        return null;
    }

    private boolean hasTypedData(Row row) {
        for (Cell cell : row) {
            CellType type = cell.getCellType();
            if (type == CellType.FORMULA) {
                try {
                    type = FormulaCache.read(cell).getCellType();
                } catch (FormulaCacheException e) {
                    // 缺失缓存无法作为判定依据，实际合并时会检查需要写入的公式。
                    continue;
                }
            }
            if (type == CellType.NUMERIC || type == CellType.BOOLEAN || type == CellType.ERROR) return true;
        }
        return false;
    }

    private SourceMapping buildSourceMapping(List<String> sourceHeaders, List<String> standardHeaders,
                                             String location) throws IOException {
        Map<String, Integer> sourceIndex = new HashMap<>();
        for (int i = 0; i < sourceHeaders.size(); i++) {
            ExcelUtils.putUniqueHeader(sourceIndex, normalizeHeader(sourceHeaders.get(i)), i, location);
        }

        Map<Integer, Integer> targetToSourceIndex = new LinkedHashMap<>();
        List<String> missingHeaders = new ArrayList<>();
        for (int i = 0; i < standardHeaders.size(); i++) {
            Integer sourceIdx = sourceIndex.get(normalizeHeader(standardHeaders.get(i)));
            if (sourceIdx == null) {
                missingHeaders.add(standardHeaders.get(i));
            } else {
                targetToSourceIndex.put(i, sourceIdx);
            }
        }
        return new SourceMapping(targetToSourceIndex, missingHeaders);
    }

    private List<String> readHeaders(Row headerRow) throws FormulaCacheException {
        List<String> headers = new ArrayList<>();
        if (headerRow == null || headerRow.getLastCellNum() <= 0) {
            return headers;
        }
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            String header = readMergeCellText(headerRow.getCell(i));
            if (header != null) {
                header = header.trim();
            }
            headers.add(header == null ? "" : header);
        }
        trimTrailingEmptyHeaders(headers);
        boolean hasHeader = false;
        for (String header : headers) {
            if (header != null && !header.trim().isEmpty()) {
                hasHeader = true;
                break;
            }
        }
        if (!hasHeader) {
            headers.clear();
        }
        return headers;
    }

    private void trimTrailingEmptyHeaders(List<String> headers) {
        for (int i = headers.size() - 1; i >= 0; i--) {
            String header = headers.get(i);
            if (header != null && !header.trim().isEmpty()) {
                return;
            }
            headers.remove(i);
        }
    }

    private static boolean hasAnyCellValue(Row row) {
        if (row == null || row.getLastCellNum() <= 0) {
            return false;
        }
        for (Cell cell : row) {
            // 已保存的错误值和公式的空文本结果也属于数据。
            if (cell.getCellType() == CellType.ERROR || cell.getCellType() == CellType.FORMULA) {
                return true;
            }
            String value = ExcelUtils.getCellValueAsString(cell);
            if (value != null && !value.trim().isEmpty()) {
                return true;
            }
        }
        return false;
    }

    private static void widenFromSource(Sheet source, int sourceColumn, Sheet target, int targetColumn) {
        int width = Math.min(source.getColumnWidth(sourceColumn), 255 * 256);
        if (width > target.getColumnWidth(targetColumn)) target.setColumnWidth(targetColumn, width);
    }

    private static String readMergeCellText(Cell cell) throws FormulaCacheException {
        if (cell != null && cell.getCellType() == CellType.ERROR) {
            return FormulaError.forInt(cell.getErrorCellValue()).getString();
        }
        if (cell != null && cell.getCellType() == CellType.FORMULA) {
            CellValue result = FormulaCache.read(cell);
            if (result.getCellType() == CellType.ERROR) {
                return FormulaError.forInt(result.getErrorValue()).getString();
            }
        }
        return ExcelUtils.getCellValueAsString(cell);
    }

    private static void copyMergedCellValue(Cell source, Cell target, Map<String, MergedNumberStyle> styles)
            throws FormulaCacheException {
        if (source.getCellType() == CellType.NUMERIC) {
            writeMergedNumber(source, target, source.getNumericCellValue(), styles);
            return;
        }
        if (source.getCellType() != CellType.FORMULA) {
            ExcelUtils.copyCellValue(source, target);
            return;
        }
        CellValue value = FormulaCache.read(source);
        switch (value.getCellType()) {
            case NUMERIC:
                writeMergedNumber(source, target, value.getNumberValue(), styles);
                break;
            case STRING:
                target.setCellValue(value.getStringValue());
                break;
            case BOOLEAN:
                target.setCellValue(value.getBooleanValue());
                break;
            case ERROR:
                target.setCellErrorValue(value.getErrorValue());
                break;
            default:
                throw new FormulaCacheException(source, null);
        }
    }

    private static void writeMergedNumber(Cell source, Cell target, double number, Map<String, MergedNumberStyle> styles) {
        CellStyle original = source.getCellStyle();
        String format = original.getDataFormatString();
        if (format != null && !"General".equalsIgnoreCase(format)) {
            MergedNumberStyle style = styles.computeIfAbsent(format,
                    code -> new MergedNumberStyle(target.getSheet().getWorkbook(), original.getDataFormat(), code));
            target.setCellStyle(style.output);
            if (style.calendarDate) {
                Workbook input = source.getSheet().getWorkbook();
                boolean from1904 = input instanceof Date1904Support ? ((Date1904Support) input).isDate1904()
                        : ((HSSFWorkbook) input).getWorkbook().isUsing1904DateWindowing();
                // 合并结果使用 1900 日期系统。直接平移天数可保留小数精度，也不涉及本机时区。
                if (from1904) number += 1462;
            }
        }
        target.setCellValue(number);
    }

    private static final class MergedNumberStyle {
        private static final Pattern LITERALS = Pattern.compile(
                "\"(?:[^\"\\\\]|\\\\.)*\"|\\\\.|_.|\\*.|\\[(?![hms]+\\])[^]]*]|am/pm|a/p", Pattern.CASE_INSENSITIVE);
        private static final Pattern DATE_FIELDS = Pattern.compile("\\[[hms]+\\]|y+|d+|m+|h+|s+");
        final CellStyle output;
        final boolean calendarDate;

        MergedNumberStyle(Workbook workbook, short inputFormatIndex, String code) {
            output = workbook.createCellStyle();
            output.setDataFormat(workbook.createDataFormat().getFormat(code));
            calendarDate = isCalendarDate(inputFormatIndex, code);
        }

        private static boolean isCalendarDate(short index, String code) {
            String fields = LITERALS.matcher(code).replaceAll("").toLowerCase(Locale.ROOT);
            int end = fields.indexOf(';');
            if (end >= 0) fields = fields.substring(0, end);
            if (!DateUtil.isADateFormat(index, fields)) return false;
            // m/mm 在小时之后或秒之前表示分钟；[h]、[m]、[s] 都表示经过的时长。
            Matcher tokens = DATE_FIELDS.matcher(fields);
            boolean afterHour = false;
            boolean possibleMonth = false;
            while (tokens.find()) {
                String token = tokens.group();
                boolean elapsed = token.charAt(0) == '[';
                char field = token.charAt(elapsed ? 1 : 0);
                if (possibleMonth && field != 's') return true;
                if (field == 'y' || field == 'd' || (!elapsed && field == 'm' && token.length() >= 3)) return true;
                possibleMonth = !elapsed && field == 'm' && !afterHour;
                afterHour = field == 'h';
            }
            return possibleMonth;
        }
    }

    private File prepareOutputFile(String outputFilePath) throws IOException {
        if (outputFilePath == null || outputFilePath.trim().isEmpty()) {
            throw new IOException("请选择输出文件");
        }
        File outputFile = new File(outputFilePath);
        File outputDir = outputFile.getParentFile();
        if (outputDir != null && !outputDir.exists() && !outputDir.mkdirs()) {
            throw new IOException("无法创建输出目录: " + outputDir.getAbsolutePath());
        }
        if (outputDir != null && !outputDir.canWrite()) {
            throw new IOException("无法写入输出目录: " + outputDir.getAbsolutePath());
        }
        return outputFile;
    }

    private void validateReadableFile(File file, String label) throws IOException {
        if (file == null || !file.exists() || !file.canRead()) {
            throw new IOException("无法读取" + label + ": " + (file == null ? "" : file.getAbsolutePath()));
        }
    }

    private void notifyProgress(ProgressListener progressListener, int current, int total) {
        if (progressListener != null) {
            progressListener.onProgress(current, Math.max(total, 1));
        }
    }

    private static String normalizeHeader(String header) {
        if (header == null) {
            return "";
        }
        String cleaned = ExcelUtils.cleanString(header);
        return cleaned == null ? "" : cleaned.toLowerCase(Locale.ROOT);
    }

    private static String fileName(File file) {
        return file == null ? "" : file.getName();
    }

    private static String columnName(int oneBasedColumnIndex) {
        StringBuilder name = new StringBuilder();
        int col = oneBasedColumnIndex;
        while (col > 0) {
            col--;
            name.insert(0, (char) ('A' + (col % 26)));
            col /= 26;
        }
        return name.toString();
    }

    public interface ProgressListener {
        void onProgress(int current, int total);
    }

    private static class SourceMapping {
        private final Map<Integer, Integer> targetToSourceIndex;
        private final List<String> missingHeaders;

        private SourceMapping(Map<Integer, Integer> targetToSourceIndex, List<String> missingHeaders) {
            this.targetToSourceIndex = targetToSourceIndex;
            this.missingHeaders = missingHeaders;
        }
    }

    private static class HeaderMergeWriter {
        private final SXSSFWorkbook workbook;
        private final Map<String, MergedNumberStyle> numberStyles = new HashMap<>();
        private final Sheet dataSheet;
        private final Sheet reportSheet;
        private final List<HeaderReportEntry> reports = new ArrayList<>();
        private List<String> standardHeaders;
        private int dataRowIndex = 1;

        private HeaderMergeWriter(SXSSFWorkbook workbook) {
            this.workbook = workbook;
            this.dataSheet = workbook.createSheet(HEADER_DATA_SHEET_NAME);
            this.reportSheet = workbook.createSheet(HEADER_REPORT_SHEET_NAME);
        }

        private void ensureStandardHeaders(List<String> headers) {
            if (standardHeaders != null) {
                return;
            }
            standardHeaders = new ArrayList<>(headers);
            Row headerRow = dataSheet.createRow(0);
            for (int i = 0; i < standardHeaders.size(); i++) {
                headerRow.createCell(i).setCellValue(standardHeaders.get(i));
            }
        }

        private List<String> getStandardHeaders() {
            return standardHeaders == null ? new ArrayList<>() : new ArrayList<>(standardHeaders);
        }

        private void appendDataRow(Row sourceRow, Map<Integer, Integer> targetToSourceIndex)
                throws FormulaCacheException {
            Row outputRow = dataSheet.createRow(dataRowIndex++);
            if (standardHeaders == null) {
                return;
            }
            for (int targetIndex = 0; targetIndex < standardHeaders.size(); targetIndex++) {
                Cell outputCell = outputRow.createCell(targetIndex);
                Integer sourceIndex = targetToSourceIndex.get(targetIndex);
                if (sourceIndex == null) {
                    outputCell.setCellValue("");
                    continue;
                }
                Cell sourceCell = sourceRow.getCell(sourceIndex);
                if (sourceCell != null) {
                    copyMergedCellValue(sourceCell, outputCell, numberStyles);
                }
            }
        }

        private void addReport(String sourceFile, String sheetName, String status,
                               int readRows, int mergedRows, String remark) {
            reports.add(new HeaderReportEntry(sourceFile, sheetName, status, readRows, mergedRows, remark));
        }

        private MergeSummary finish() throws IOException {
            int successCount = 0;
            int partialCount = 0;
            String firstFailureReason = null;
            for (HeaderReportEntry report : reports) {
                if ("成功".equals(report.status)) {
                    successCount++;
                } else if ("部分成功".equals(report.status)) {
                    partialCount++;
                } else if (firstFailureReason == null && report.remark != null && !report.remark.isEmpty()) {
                    firstFailureReason = report.remark;
                }
            }
            if (successCount + partialCount == 0) {
                String reason = firstFailureReason == null ? "未找到有效表头" : firstFailureReason;
                throw new IOException("没有成功合并的工作表: " + reason);
            }
            MergeSummary summary = new MergeSummary(successCount, partialCount, reports.size() - successCount - partialCount);
            writeReport(summary);
            return summary;
        }

        private void writeReport(MergeSummary summary) {
            Row header = reportSheet.createRow(0);
            String[] headers = {"来源文件", "来源Sheet", "状态", "读取数据行数", "合并行数", "备注"};
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i).setCellValue(headers[i]);
            }

            int rowIndex = 1;
            int totalReadRows = 0;
            int totalMergedRows = 0;
            for (HeaderReportEntry report : reports) {
                Row row = reportSheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(report.sourceFile);
                row.createCell(1).setCellValue(report.sheetName);
                row.createCell(2).setCellValue(report.status);
                row.createCell(3).setCellValue(report.readRows);
                row.createCell(4).setCellValue(report.mergedRows);
                row.createCell(5).setCellValue(report.remark == null ? "" : report.remark);
                totalReadRows += report.readRows;
                totalMergedRows += report.mergedRows;
            }

            Row total = reportSheet.createRow(rowIndex);
            total.createCell(0).setCellValue("总计");
            total.createCell(2).setCellValue(summary.formatSummary());
            total.createCell(3).setCellValue(totalReadRows);
            total.createCell(4).setCellValue(totalMergedRows);
        }
    }

    private static class WpsMergeWriter {
        private final SXSSFWorkbook workbook;
        private final Map<String, MergedNumberStyle> numberStyles = new HashMap<>();
        private final Sheet reportSheet;
        private final Sheet dataSheet;
        private final List<WpsReportEntry> reports = new ArrayList<>();
        private int dataRowIndex = 0;
        private int successCount = 0;
        private int failCount = 0;
        private int mergedRows = 0;

        private WpsMergeWriter(SXSSFWorkbook workbook) {
            this.workbook = workbook;
            this.reportSheet = workbook.createSheet(WPS_REPORT_SHEET_NAME);
            this.dataSheet = workbook.createSheet(WPS_DATA_SHEET_NAME);
        }

        private void processSheet(File inputFile, int workbookIndex, Sheet sourceSheet) throws FormulaCacheException {
            String sheetName = sourceSheet == null ? "" : sourceSheet.getSheetName();
            if (sourceSheet == null) {
                addFailure(inputFile, workbookIndex, sheetName, "失败：没有数据");
                return;
            }
            List<Row> rows = nonEmptyRows(sourceSheet);
            if (rows.isEmpty()) {
                addFailure(inputFile, workbookIndex, sheetName, "失败：没有数据");
                return;
            }

            int startRow = dataRowIndex + 1;
            int maxSourceColumn = maxSourceColumn(rows);
            for (int i = 0; i < maxSourceColumn; i++) {
                widenFromSource(sourceSheet, i, dataSheet, i + 1);
            }
            for (Row sourceRow : rows) {
                Row outputRow = dataSheet.createRow(dataRowIndex++);
                outputRow.createCell(0).setCellValue("[" + workbookIndex + "]" + sheetName);
                for (int i = 0; i < maxSourceColumn; i++) {
                    Cell sourceCell = sourceRow.getCell(i);
                    Cell outputCell = outputRow.createCell(i + 1);
                    if (sourceCell != null) {
                        copyMergedCellValue(sourceCell, outputCell, numberStyles);
                    }
                }
            }
            int endRow = dataRowIndex;
            successCount++;
            mergedRows += rows.size();
            String range = WPS_DATA_SHEET_NAME + "!B" + startRow + ":" + columnName(maxSourceColumn + 1) + endRow;
            reports.add(new WpsReportEntry(fileName(inputFile), sheetName, "成功", range));
        }

        private void addFailure(File inputFile, int workbookIndex, String sheetName, String status) {
            failCount++;
            reports.add(new WpsReportEntry(fileName(inputFile), sheetName, status, ""));
        }

        private MergeSummary finish() throws IOException {
            if (successCount == 0) {
                String reason = reports.isEmpty() ? "没有数据" : reports.get(0).status;
                throw new IOException("没有成功合并的工作表: " + reason);
            }
            Row summary = reportSheet.createRow(1);
            summary.createCell(1).setCellValue("报告：成功合并 " + successCount + " 个工作表，失败 " + failCount +
                    " 个，共 " + mergedRows + " 行数据。");

            Row header = reportSheet.createRow(3);
            header.createCell(1).setCellValue("工作簿");
            header.createCell(2).setCellValue("工作表");
            header.createCell(3).setCellValue("合并状态");
            header.createCell(4).setCellValue("合并后的位置");

            int rowIndex = 4;
            for (WpsReportEntry report : reports) {
                Row row = reportSheet.createRow(rowIndex++);
                row.createCell(1).setCellValue(report.workbookName);
                row.createCell(2).setCellValue(report.sheetName);
                row.createCell(3).setCellValue(report.status);
                row.createCell(4).setCellValue(report.location);
            }
            return new MergeSummary(successCount, 0, failCount);
        }

        private List<Row> nonEmptyRows(Sheet sheet) {
            List<Row> rows = new ArrayList<>();
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (hasAnyCellValue(row)) {
                    rows.add(row);
                }
            }
            return rows;
        }

        private int maxSourceColumn(List<Row> rows) {
            int max = 0;
            for (Row row : rows) {
                if (row != null && row.getLastCellNum() > max) {
                    max = row.getLastCellNum();
                }
            }
            return max;
        }
    }

    private static class HeaderReportEntry {
        private final String sourceFile;
        private final String sheetName;
        private final String status;
        private final int readRows;
        private final int mergedRows;
        private final String remark;

        private HeaderReportEntry(String sourceFile, String sheetName, String status,
                                  int readRows, int mergedRows, String remark) {
            this.sourceFile = sourceFile;
            this.sheetName = sheetName;
            this.status = status;
            this.readRows = readRows;
            this.mergedRows = mergedRows;
            this.remark = remark;
        }
    }

    private static class WpsReportEntry {
        private final String workbookName;
        private final String sheetName;
        private final String status;
        private final String location;

        private WpsReportEntry(String workbookName, String sheetName, String status, String location) {
            this.workbookName = workbookName;
            this.sheetName = sheetName;
            this.status = status;
            this.location = location;
        }
    }
}
