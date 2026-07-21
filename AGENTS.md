# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## Project Overview

Java Swing desktop application (Excel工具集) providing Excel file processing tools:
1. **Excel Processor** - Maps headers between source and template Excel files, merging multiple source columns into one target column
2. **Header Mapping Tool** - Creates and saves mapping relationships between fields
3. **Cross-Sheet Filler** - Fills data across two Excel files by matching a key column
4. **Duplicate Column Merger** - Merges duplicate rows by key column, concatenating values with `|`

## Build & Run

```bash
mvn clean compile          # Compile
mvn package                # Build shaded JAR (target/excel-processor-1.0-SNAPSHOT.jar)
java -jar target/excel-processor-1.0-SNAPSHOT.jar   # Run GUI
```

Launch4j wraps the JAR into `ExcelProcessor.exe` via `launch4j-config.xml`.

No automated tests exist. Testing is manual with sample files in `test-data/` (`测试A表.xlsx`, `测试B表.xlsx`).

## Architecture

**Entry point:** `com.xzy.ExcelTools.main()` — 5-tab Swing JFrame (1000x700).

All long operations use `SwingWorker` off the EDT with modal progress dialogs.

| Package | Key Classes | Role |
|---------|------------|------|
| `com.xzy` | `ExcelTools` (2074 lines) | Main app: Excel Processor, Build Mapping, Apply Mapping, Cross-Sheet, Duplicate-Merge tabs |
| `com.xzy` | `CrossSheetFiller` @Deprecated | Cross-sheet key matching + value filling, used by ExcelTools tab |
| `com.xzy` | `DuplicateColumnMerger` @Deprecated | Row dedup by key column, used by ExcelTools tab |

Legacy code (not compiled, kept for reference): `src/legacy/`

### Excel Processor Data Flow

1. Parse source + template headers (SAX for `.xlsx`, POI for `.xls`)
2. Build `HeaderMapping` list — multiple source headers can map to one target header
3. Stream source rows → write to target via `SXSSFWorkbook` (100-row memory window, 1000-row batch flush)
4. Multi-source merge: `COMMA_SEPARATED` for "手机号码" field (comma, no header names); `PIPE_SEPARATED` for others (`源表头：值|源表头：值`)

### CrossSheetFiller Algorithm

- Index target table rows into `Map<String, List<Row>>` by key column
- Iterate source table, look up key (case-insensitive, whitespace-trimmed via `cleanString()`), write value to all matching target rows

### DuplicateColumnMerger Algorithm

- Group rows by key column into `Map<String, List<List<String>>>`
- Merge columns: join unique values with `|` using `LinkedHashSet`
- Non-merge columns: take first unique value

## Key Technical Details

- **Java target:** 1.8 (source and target)
- **Dependencies:** Apache POI 5.2.3, Log4j 2.20.0
- **Memory:** SXSSF with 100-row window; `IOUtils.setByteArrayMaxOverride(1000000000)` (1GB); max file size 1000MB
- **Logging:** Log4j2 via `src/main/resources/log4j2.xml`, console appender at INFO level
- **Packaging:** maven-shade-plugin 3.2.4 creates fat JAR with `com.xzy.ExcelTools` as main class
- **GUI language:** Chinese (Simplified)
