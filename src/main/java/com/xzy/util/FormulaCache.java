package com.xzy.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.io.File;
import java.io.IOException;

/** 读取 POI 公式单元格已保存的结果，不重新计算公式。 */
public final class FormulaCache {
    private FormulaCache() {
    }

    public static CellValue read(Cell cell) throws FormulaCacheException {
        try {
            CellType resultType = cell.getCachedFormulaResultType();
            if (cell instanceof XSSFCell) {
                // XSSF 会把缺失的数值缓存读为 0；已保存的数值 0 和空文本缓存均合法。
                String raw = ((XSSFCell) cell).getRawValue();
                if (raw == null || (resultType != CellType.STRING && raw.trim().isEmpty())) {
                    throw new FormulaCacheException(cell, null);
                }
                if (resultType == CellType.BOOLEAN && !"0".equals(raw) && !"1".equals(raw)) {
                    throw new FormulaCacheException(cell, null);
                }
            }
            switch (resultType) {
                case NUMERIC:
                    double value = cell.getNumericCellValue();
                    if (Double.isNaN(value) || Double.isInfinite(value)) {
                        throw new FormulaCacheException(cell, null);
                    }
                    return new CellValue(value);
                case STRING:
                    return new CellValue(cell.getStringCellValue());
                case BOOLEAN:
                    return CellValue.valueOf(cell.getBooleanCellValue());
                case ERROR:
                    return CellValue.getError(cell.getErrorCellValue());
                default:
                    throw new FormulaCacheException(cell, null);
            }
        } catch (RuntimeException e) {
            throw new FormulaCacheException(cell, e);
        }
    }

    public static final class FormulaCacheException extends IOException {
        private static final long serialVersionUID = 1L;

        public FormulaCacheException(Cell cell, Throwable cause) {
            super("工作表“" + cell.getSheet().getSheetName() + "”的单元格 "
                    + cell.getAddress().formatAsString()
                    + " 缺少有效的公式缓存结果，请在 Excel 或 WPS 中重新计算并保存源文件后重试。", cause);
        }

        public IOException inFile(File source) {
            return new IOException(source.getAbsolutePath() + "：" + getMessage(), this);
        }
    }
}
