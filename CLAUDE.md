# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Java-based Excel processing tool (Excel工具集) that provides a comprehensive suite for Excel file manipulation including header mapping, data merging, cross-sheet key filling, and duplicate column merging. The application uses Java Swing for the GUI and Apache POI for Excel file processing.

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
- `com.xzy.ExcelTools` - Main application class with integrated GUI and functionality
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

### Memory Management
- Uses SXSSF (Streaming Workbook) for large Excel files
- Configured memory window size of 100 rows
- Batch processing of 1000 rows at a time
- Maximum file size limit of 1000MB per file
- Optimized for handling millions of rows

### Key Design Patterns
- **Multi-threading**: SwingWorker for background processing with progress updates
- **Event-driven**: Comprehensive GUI event handling for file operations
- **Factory Pattern**: HeaderMapping class with different format types
- **Strategy Pattern**: Different processing strategies for .xlsx vs .xls files

## Development Notes

### File Processing Approach
- Uses SAX parsing for reading headers to minimize memory usage
- Supports both .xlsx and .xls formats
- Implements streaming processing for large datasets
- Progress tracking and user feedback throughout operations

### GUI Architecture  
- Tabbed interface with separate panels for different functionalities
- Real-time progress indicators and logging
- File validation and error handling with user-friendly messages
- Consistent styling and responsive layout

### Data Mapping System
- Flexible header mapping with support for many-to-one relationships
- Different merge formats (comma-separated vs pipe-separated with labels)
- Persistent mapping relationships that can be saved and reused
- Visual mapping interface with drag-and-drop style selection