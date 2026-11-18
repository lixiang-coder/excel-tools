package com.xzy;

import com.xzy.util.ExcelUtils;
import com.xzy.util.FormulaCache;
import com.xzy.util.FormulaCache.FormulaCacheException;
import com.xzy.util.SafeExport;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
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
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

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
        File inputFile = new File(inputFilePath);
        validateReadableFile(inputFile, "输入文件");
        File outputFile = prepareOutputFile(outputFilePath);

        try (SafeExport export = SafeExport.begin(outputFile, inputFile)) {
            try (FileInputStream fis = new FileInputStream(inputFile);
                 Workbook inputWorkbook = WorkbookFactory.create(fis);
                 SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(MEMORY_WINDOW_SIZE)) {
                try {
                    MergeStrategy effectiveStrategy = resolveStrategy(strategy, collectSheetSamples(inputWorkbook));
                    logger.info("表格合并策略: {}", effectiveStrategy);
                    if (effectiveStrategy == MergeStrategy.WPS_STYLE_NO_HEADER) {
                        mergeWorkbookSheetsWpsStyle(inputFile, inputWorkbook, outputWorkbook, progressListener);
                    } else {
                        mergeWorkbookSheetsHeaderBased(inputFile, inputWorkbook, outputWorkbook, progressListener);
                    }
                    try (FileOutputStream fos = new FileOutputStream(export.temporaryFile())) {
                        outputWorkbook.write(fos);
                    }
                } finally {
                    outputWorkbook.dispose();
                }
            }
            export.commit();
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
        if (inputFiles == null || inputFiles.isEmpty()) {
            throw new IOException("请选择要合并的Excel文件");
        }
        File outputFile = prepareOutputFile(outputFilePath);

        try (SafeExport export = SafeExport.begin(outputFile, inputFiles.toArray(new File[0]))) {
            MergeStrategy effectiveStrategy = resolveStrategy(strategy, collectFileSamples(inputFiles));
            logger.info("表格合并策略: {}", effectiveStrategy);
            if (effectiveStrategy == MergeStrategy.WPS_STYLE_NO_HEADER) {
                mergeFilesWpsStyle(inputFiles, export.temporaryFile(), progressListener);
            } else {
                mergeFilesHeaderBased(inputFiles, export.temporaryFile(), progressListener);
            }
            export.commit();
        }
    }

    private void mergeWorkbookSheetsHeaderBased(File inputFile, Workbook inputWorkbook, SXSSFWorkbook outputWorkbook,
                                                ProgressListener progressListener) throws IOException {
        HeaderMergeWriter writer = new HeaderMergeWriter(outputWorkbook);
        int total = inputWorkbook.getNumberOfSheets();
        for (int i = 0; i < total; i++) {
            processHeaderSheet(inputFile.getAbsolutePath(), inputWorkbook.getSheetAt(i), writer);
            notifyProgress(progressListener, i + 1, total);
        }
        writer.finish();
    }

    private void mergeFilesHeaderBased(List<File> inputFiles, File outputFile,
                                       ProgressListener progressListener) throws IOException {
        try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(MEMORY_WINDOW_SIZE)) {
            try {
                HeaderMergeWriter writer = new HeaderMergeWriter(outputWorkbook);
                int total = inputFiles.size();
                for (int i = 0; i < total; i++) {
                    processFirstSheetHeaderBased(inputFiles.get(i), writer);
                    notifyProgress(progressListener, i + 1, total);
                }
                writer.finish();
                try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                    outputWorkbook.write(fos);
                }
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

        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = WorkbookFactory.create(fis)) {
            if (workbook.getNumberOfSheets() == 0) {
                writer.addReport(inputFile.getAbsolutePath(), "", "失败", 0, 0, "未找到有效表头");
                return;
            }
            processHeaderSheet(inputFile.getAbsolutePath(), workbook.getSheetAt(0), writer);
        } catch (FormulaCacheException e) {
            throw e.inFile(inputFile);
        } catch (Exception e) {
            logger.warn("处理文件失败: {}", inputFile.getAbsolutePath(), e);
            writer.addReport(inputFile.getAbsolutePath(), "", "失败", 0, 0, e.getMessage());
        }
    }

    private void processHeaderSheet(String sourceFile, Sheet sourceSheet, HeaderMergeWriter writer)
            throws FormulaCacheException {
        String sheetName = sourceSheet == null ? "" : sourceSheet.getSheetName();
        if (sourceSheet == null) {
            writer.addReport(sourceFile, sheetName, "失败", 0, 0, "未找到有效表头");
            return;
        }

        Row headerRow = sourceSheet.getRow(0);
        List<String> headers = readHeaders(headerRow);
        if (headers.isEmpty()) {
            writer.addReport(sourceFile, sheetName, "失败", 0, 0, "未找到有效表头");
            return;
        }

        writer.ensureStandardHeaders(headers);
        SourceMapping mapping = buildSourceMapping(headers, writer.getStandardHeaders());
        int readRows = 0;
        int mergedRows = 0;
        for (int i = 1; i <= sourceSheet.getLastRowNum(); i++) {
            Row row = sourceSheet.getRow(i);
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

    private void mergeWorkbookSheetsWpsStyle(File inputFile, Workbook inputWorkbook, SXSSFWorkbook outputWorkbook,
                                             ProgressListener progressListener) throws IOException {
        WpsMergeWriter writer = new WpsMergeWriter(outputWorkbook);
        int total = inputWorkbook.getNumberOfSheets();
        for (int i = 0; i < total; i++) {
            writer.processSheet(inputFile, 1, inputWorkbook.getSheetAt(i));
            notifyProgress(progressListener, i + 1, total);
        }
        writer.finish();
    }

    private void mergeFilesWpsStyle(List<File> inputFiles, File outputFile,
                                    ProgressListener progressListener) throws IOException {
        try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(MEMORY_WINDOW_SIZE)) {
            try {
                WpsMergeWriter writer = new WpsMergeWriter(outputWorkbook);
                int total = inputFiles.size();
                for (int i = 0; i < total; i++) {
                    File inputFile = inputFiles.get(i);
                    if (inputFile == null || !inputFile.exists() || !inputFile.canRead()) {
                        writer.addFailure(inputFile, i + 1, "", "失败：无法读取文件");
                    } else {
                        try (FileInputStream fis = new FileInputStream(inputFile);
                             Workbook workbook = WorkbookFactory.create(fis)) {
                            if (workbook.getNumberOfSheets() == 0) {
                                writer.addFailure(inputFile, i + 1, "", "失败：没有数据");
                            } else {
                                writer.processSheet(inputFile, i + 1, workbook.getSheetAt(0));
                            }
                        } catch (FormulaCacheException e) {
                            throw e.inFile(inputFile);
                        } catch (Exception e) {
                            logger.warn("处理文件失败: {}", inputFile.getAbsolutePath(), e);
                            writer.addFailure(inputFile, i + 1, "", "失败：" + e.getMessage());
                        }
                    }
                    notifyProgress(progressListener, i + 1, total);
                }
                writer.finish();
                try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                    outputWorkbook.write(fos);
                }
            } finally {
                outputWorkbook.dispose();
            }
        }
    }

    private MergeStrategy resolveStrategy(MergeStrategy requestedStrategy, List<List<String>> samples) {
        if (requestedStrategy == null || requestedStrategy == MergeStrategy.AUTO) {
            return looksLikeHeaderRows(samples) ? MergeStrategy.HEADER_BASED : MergeStrategy.WPS_STYLE_NO_HEADER;
        }
        return requestedStrategy;
    }

    private List<List<String>> collectSheetSamples(Workbook workbook) throws FormulaCacheException {
        List<List<String>> samples = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            List<String> row = firstNonEmptyRowValues(sheet);
            if (!row.isEmpty()) {
                samples.add(row);
            }
        }
        return samples;
    }

    private List<List<String>> collectFileSamples(List<File> inputFiles) throws IOException {
        List<List<String>> samples = new ArrayList<>();
        for (File inputFile : inputFiles) {
            if (inputFile == null || !inputFile.exists() || !inputFile.canRead()) {
                continue;
            }
            try (FileInputStream fis = new FileInputStream(inputFile);
                 Workbook workbook = WorkbookFactory.create(fis)) {
                if (workbook.getNumberOfSheets() > 0) {
                    List<String> row = firstNonEmptyRowValues(workbook.getSheetAt(0));
                    if (!row.isEmpty()) {
                        samples.add(row);
                    }
                }
            } catch (FormulaCacheException e) {
                throw e.inFile(inputFile);
            } catch (Exception e) {
                logger.warn("采集合并样本失败: {}", inputFile.getAbsolutePath(), e);
            }
        }
        return samples;
    }

    private boolean looksLikeHeaderRows(List<List<String>> samples) {
        if (samples.isEmpty()) {
            return true;
        }
        if (samples.size() == 1) {
            return rowLooksLikeHeader(samples.get(0));
        }

        List<String> base = normalizedNonEmptyValues(samples.get(0));
        if (base.isEmpty()) {
            return false;
        }
        double overlapSum = 0;
        int compared = 0;
        for (int i = 1; i < samples.size(); i++) {
            List<String> current = normalizedNonEmptyValues(samples.get(i));
            if (current.isEmpty()) {
                continue;
            }
            int hits = 0;
            for (String value : current) {
                if (base.contains(value)) {
                    hits++;
                }
            }
            overlapSum += hits * 1.0 / Math.min(base.size(), current.size());
            compared++;
        }
        if (compared > 0 && overlapSum / compared >= 0.6) {
            return true;
        }
        int headerLikeRows = 0;
        for (List<String> sample : samples) {
            if (rowLooksLikeHeader(sample)) {
                headerLikeRows++;
            }
        }
        return headerLikeRows >= Math.ceil(samples.size() * 0.8);
    }

    private boolean rowLooksLikeHeader(List<String> row) {
        int nonEmpty = 0;
        int dataLike = 0;
        for (String value : row) {
            if (value == null || value.trim().isEmpty()) {
                continue;
            }
            nonEmpty++;
            if (looksLikeDataValue(value)) {
                dataLike++;
            }
        }
        return nonEmpty > 0 && dataLike == 0;
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

    private List<String> firstNonEmptyRowValues(Sheet sheet) throws FormulaCacheException {
        List<String> values = new ArrayList<>();
        if (sheet == null) {
            return values;
        }
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (hasAnyCellValue(row)) {
                int lastCell = Math.max(row.getLastCellNum(), (short) 0);
                for (int c = 0; c < lastCell; c++) {
                    values.add(readMergeCellText(row.getCell(c)));
                }
                trimTrailingEmptyHeaders(values);
                return values;
            }
        }
        return values;
    }

    private SourceMapping buildSourceMapping(List<String> sourceHeaders, List<String> standardHeaders) {
        Map<String, Integer> sourceIndex = new HashMap<>();
        for (int i = 0; i < sourceHeaders.size(); i++) {
            String normalized = normalizeHeader(sourceHeaders.get(i));
            if (!normalized.isEmpty() && !sourceIndex.containsKey(normalized)) {
                sourceIndex.put(normalized, i);
            }
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

    private List<String> normalizedNonEmptyValues(List<String> values) {
        List<String> normalized = new ArrayList<>();
        for (String value : values) {
            String item = normalizeHeader(value);
            if (!item.isEmpty()) {
                normalized.add(item);
            }
        }
        return normalized;
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
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            // 公式的空文本和错误结果也必须保留，不能把整行当作空行跳过。
            if (cell != null && cell.getCellType() == CellType.FORMULA) {
                return true;
            }
            String value = ExcelUtils.getCellValueAsString(cell);
            if (value != null && !value.trim().isEmpty()) {
                return true;
            }
        }
        return false;
    }

    private static String readMergeCellText(Cell cell) throws FormulaCacheException {
        if (cell != null && cell.getCellType() == CellType.FORMULA) {
            CellValue result = FormulaCache.read(cell);
            if (result.getCellType() == CellType.ERROR) {
                return FormulaError.forInt(result.getErrorValue()).getString();
            }
        }
        return ExcelUtils.getCellValueAsString(cell);
    }

    private static void copyMergedCellValue(Cell source, Cell target) throws FormulaCacheException {
        if (source.getCellType() != CellType.FORMULA) {
            ExcelUtils.copyCellValue(source, target);
            return;
        }
        CellValue value = FormulaCache.read(source);
        switch (value.getCellType()) {
            case NUMERIC:
                target.setCellValue(value.getNumberValue());
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
                    copyMergedCellValue(sourceCell, outputCell);
                }
            }
        }

        private void addReport(String sourceFile, String sheetName, String status,
                               int readRows, int mergedRows, String remark) {
            reports.add(new HeaderReportEntry(sourceFile, sheetName, status, readRows, mergedRows, remark));
        }

        private void finish() throws IOException {
            boolean hasSuccess = false;
            String firstFailureReason = null;
            for (HeaderReportEntry report : reports) {
                if ("成功".equals(report.status) || "部分成功".equals(report.status)) {
                    hasSuccess = true;
                } else if (firstFailureReason == null && report.remark != null && !report.remark.isEmpty()) {
                    firstFailureReason = report.remark;
                }
            }
            if (!hasSuccess) {
                String reason = firstFailureReason == null ? "未找到有效表头" : firstFailureReason;
                throw new IOException("没有成功合并的工作表: " + reason);
            }
            writeReport();
        }

        private void writeReport() {
            Row header = reportSheet.createRow(0);
            String[] headers = {"来源文件", "来源Sheet", "状态", "读取数据行数", "合并行数", "备注"};
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i).setCellValue(headers[i]);
            }

            int rowIndex = 1;
            int totalReadRows = 0;
            int totalMergedRows = 0;
            int successCount = 0;
            int failCount = 0;
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
                if ("失败".equals(report.status)) {
                    failCount++;
                } else {
                    successCount++;
                }
            }

            Row total = reportSheet.createRow(rowIndex);
            total.createCell(0).setCellValue("总计");
            total.createCell(2).setCellValue("成功 " + successCount + " / 失败 " + failCount);
            total.createCell(3).setCellValue(totalReadRows);
            total.createCell(4).setCellValue(totalMergedRows);
        }
    }

    private static class WpsMergeWriter {
        private final SXSSFWorkbook workbook;
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
            for (Row sourceRow : rows) {
                Row outputRow = dataSheet.createRow(dataRowIndex++);
                outputRow.createCell(0).setCellValue("[" + workbookIndex + "]" + sheetName);
                for (int i = 0; i < maxSourceColumn; i++) {
                    Cell sourceCell = sourceRow.getCell(i);
                    Cell outputCell = outputRow.createCell(i + 1);
                    if (sourceCell != null) {
                        copyMergedCellValue(sourceCell, outputCell);
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

        private void finish() throws IOException {
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
