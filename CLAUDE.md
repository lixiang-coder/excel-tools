# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Java-based Excel processing tool (Excel工具集) that provides a comprehensive suite for Excel file manipulation including header mapping, data merging, cross-sheet key filling, and duplicate column merging. The application uses Java Swing for the GUI and Apache POI for Excel file processing.

**Recently optimized for low-memory systems** to handle 100,000+ row Excel files without OOM errors.

## Common Development Commands

### Building the Project
```bash
# Build the project with dependencies
mvn clean compile

# Create executable JAR with all dependencies
mvn clean package

# Create Windows executable (requires Launch4j installed)
# The launch4j-config.xml is configured to create ExcelProcessor.exe
```

### Running the Application

**Recommended (with memory optimization)**:
```bash
# Auto-detect system memory and apply optimal JVM settings
start_optimized.bat

# Or manually with memory-optimized parameters:
# Low memory (<1GB)
java -Xms128m -Xmx512m -XX:+UseSerialGC -jar target/excel-processor-1.0-SNAPSHOT-jar-with-dependencies.jar

# Medium memory (1-2GB)
java -Xms256m -Xmx1024m -XX:+UseG1GC -XX:MaxGCPauseMillis=200 -jar target/excel-processor-1.0-SNAPSHOT-jar-with-dependencies.jar

# High memory (>2GB)
java -Xms512m -Xmx2048m -XX:+UseG1GC -XX:MaxGCPauseMillis=100 -jar target/excel-processor-1.0-SNAPSHOT-jar-with-dependencies.jar
```

**Standard execution**:
```bash
# Run directly with Maven
mvn exec:java -Dexec.mainClass="com.xzy.ExcelTools"

# Run the compiled JAR
java -jar target/excel-processor-1.0-SNAPSHOT-jar-with-dependencies.jar

# Or run the Windows executable (after building)
./ExcelProcessor.exe
```

## Architecture Overview

### Main Classes Structure
- `com.xzy.ExcelTools` - Main application class with integrated GUI and functionality (entry point)
- `com.xzy.MemoryConfigManager` - **[NEW]** Dynamic memory configuration based on system resources
- `com.xzy.ChunkedMappingProcessor` - **[NEW]** Chunked processing for large mapping relations to avoid OOM
- `com.xzy.CrossSheetFiller` - Cross-sheet primary key filling functionality
- `com.xzy.DuplicateColumnMerger` - Duplicate column merging functionality
- `com.oldFile.*` - Legacy implementations (CustomerIdMerger, ExcelDataFillingTool, TableMappingTool, ExcelProcessor)
- `com.tools.ExcelMerger` - Additional Excel merging utilities

### Core Features
1. **Excel Processor Tab** - Main Excel processing with header mapping and data transformation
2. **Mapping Relations Tab** - Build and manage field mapping relationships
3. **Apply Mapping Tab** - Apply saved mapping relationships to new data
4. **Cross-Sheet Key Filling** - Fill primary keys across different Excel sheets
5. **Duplicate Column Merger** - Merge duplicate columns based on customer ID

### Technology Stack
- **GUI Framework**: Java Swing with modern styling and progress dialogs
- **Excel Processing**: Apache POI (5.2.3) with SXSSF for large file support
- **Logging**: Log4j2 (2.20.0)
- **Build Tool**: Maven with Shade plugin for fat JAR creation
- **Packaging**: Launch4j for Windows executable creation

### Memory Management (Optimized)

**Dynamic Configuration** - The system automatically adjusts parameters based on available memory:
- **Low memory (<512MB)**: Window 5 rows, Batch 50 rows, Max file 50MB
- **Medium memory (512MB-1GB)**: Window 20 rows, Batch 200 rows, Max file 200MB
- **High memory (>1GB)**: Window 50 rows, Batch 500 rows, Max file 500MB

**Key Optimization Features**:
- Uses SXSSF (Streaming Workbook) with dynamic window sizing
- Automatic memory monitoring with 85% threshold for GC triggers
- Chunked mapping relations stored in temp files (not all in memory)
- Batch size dynamically adjusts based on file size
- Temporary file compression enabled on low-memory systems

**Processing Limits**:
- Supports .xlsx and .xls formats
- Handles 100,000+ rows on low-spec machines
- Progress tracking and user feedback throughout operations

### Key Design Patterns
- **Multi-threading**: SwingWorker for background processing with progress updates
- **Event-driven**: Comprehensive GUI event handling for file operations
- **Factory Pattern**: HeaderMapping class with different format types
- **Strategy Pattern**: Different processing strategies for .xlsx vs .xls files
- **Singleton**: MemoryConfigManager for centralized memory configuration

## Development Notes

### File Processing Approach
- Uses SAX parsing for reading headers to minimize memory usage
- Supports both .xlsx and .xls formats
- Implements streaming processing for large datasets
- Progress tracking and user feedback throughout operations
- **Chunked processing** for mapping relations to prevent memory overflow

### GUI Architecture
- Tabbed interface with separate panels for different functionalities
- Real-time progress indicators and logging
- File validation and error handling with user-friendly messages
- Memory status display in logs
- Consistent styling and responsive layout

### Data Mapping System
- Flexible header mapping with support for many-to-one relationships
- Different merge formats (comma-separated vs pipe-separated with labels)
- Persistent mapping relationships that can be saved and reused
- Visual mapping interface with drag-and-drop style selection
- **Large mapping sets stored in temporary files to reduce memory footprint**

## Memory Optimization Implementation

### When to Use Optimized Features
- Processing files >50MB
- Working on machines with <2GB RAM
- Handling >50,000 rows of data
- Building large mapping relationships (>10,000 mappings)

### Key Components

**MemoryConfigManager**:
- Auto-detects system memory and sets optimal parameters
- Monitors real-time memory usage
- Triggers garbage collection at 85% threshold
- Provides memory status logging

**ChunkedMappingProcessor**:
- Splits large mapping relations into chunks
- Stores chunks in temporary files (auto-cleanup on exit)
- Processes mappings without loading all into memory
- Ideal for mapping relationships with >10,000 entries

### Troubleshooting OOM Issues
1. Use `start_optimized.bat` for automatic memory configuration
2. Check logs for memory status and warnings
3. Close other memory-intensive applications
4. Consider splitting very large files (>500MB) into smaller chunks
5. Ensure 64-bit Java runtime is installed for better memory management

### Temporary Files
- Location: System temp directory
- Prefix: `excel_mapping_`
- Auto-cleanup: On program exit
- Manual cleanup: Delete directories starting with `excel_mapping_` in temp folder

## Important Implementation Details

### SAX Parsing for Headers
The `HeaderHandler` inner class in `ExcelTools` uses SAX parsing to read only the header row from Excel files, avoiding loading the entire file into memory.

### SXSSF Configuration
Dynamic window size is set via `MemoryConfigManager.getMemoryWindowSize()` instead of hardcoded values. Compression is enabled automatically on low-memory systems.

### Batch Processing
Batch size is calculated dynamically based on both system memory and file size using `MemoryConfigManager.getOptimalBatchSize(fileSizeMB)`.

### Multi-Format Merging
When multiple source headers map to one target header, the `HeaderMapping.FormatType` determines the separator:
- `COMMA_SEPARATED`: Used for fields like "手机号码" (phone numbers)
- `PIPE_SEPARATED`: Default, includes source header labels

## Common Issues and Solutions

**Issue**: OOM error during processing
- **Solution**: Use `start_optimized.bat` or set appropriate JVM parameters

**Issue**: Slow processing speed
- **Solution**: Check disk space (temp files needed), use optimized launcher

**Issue**: Mapping relations taking too much memory
- **Solution**: Use `ChunkedMappingProcessor` for large mapping sets (automatically used for >10,000 mappings)

**Issue**: File too large error
- **Solution**: Message shows current max file size based on memory; increase JVM heap or split file