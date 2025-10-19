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

**推荐方式（适用于大文件和复杂数据）**:
```bash
# 自动检测系统内存并应用最优JVM设置（推荐）
start_optimized.bat

# 根据系统内存自动配置：
# - <1GB 内存: -Xmx512m（仅支持小文件）
# - 1-2GB 内存: -Xmx1024m（支持中等文件）
# - 2-4GB 内存: -Xmx2560m（支持较大文件）
# - 4-8GB 内存: -Xmx6144m（支持大文件，如250MB）
# - >8GB 内存: -Xmx8192m（支持超大文件，如500MB+）
```

**手动内存配置（用于特殊情况）**:
```bash
# 低内存 (<1GB) - 仅适用于小文件
java -Xms128m -Xmx512m -XX:+UseSerialGC -jar target/excel-processor-1.0-SNAPSHOT.jar

# 中等内存 (1-2GB) - 适用于中等文件
java -Xms256m -Xmx1024m -XX:+UseG1GC -XX:MaxGCPauseMillis=200 -jar target/excel-processor-1.0-SNAPSHOT.jar

# 高内存 (>2GB) - 适用于大文件和复杂数据
java -Xms512m -Xmx2560m -XX:+UseG1GC -XX:MaxGCPauseMillis=100 -jar target/excel-processor-1.0-SNAPSHOT.jar

# 超大文件或遇到 "allocate an array" 错误时
java -Xms1g -Xmx4g -XX:+UseG1GC -XX:MaxGCPauseMillis=100 -jar target/excel-processor-1.0-SNAPSHOT.jar
```

**标准执行（仅用于开发测试）**:
```bash
# 直接使用Maven运行
mvn exec:java -Dexec.mainClass="com.xzy.ExcelTools"

# 运行编译好的JAR（使用默认JVM设置，不推荐用于大文件）
java -jar target/excel-processor-1.0-SNAPSHOT.jar

# 或运行Windows可执行文件（编译后）
./ExcelProcessor.exe
```

## Architecture Overview

### Main Classes Structure

**核心类**:
- `com.xzy.ExcelTools` - 主应用程序类，集成 GUI 和所有功能（程序入口点）
- `com.xzy.MemoryConfigManager` - **[新增]** 基于系统资源的动态内存配置管理器
- `com.xzy.ChunkedMappingProcessor` - **[新增]** 分块处理大量映射关系以避免 OOM
- `com.xzy.StreamingExcelProcessor` - **[新增]** 流式 Excel 处理器，使用 SAX 解析大文件
- `com.xzy.StreamingMappingApplier` - **[新增]** 流式映射应用器，处理百万级数据
- `com.xzy.CrossSheetFiller` - 跨表主键数据填充功能
- `com.xzy.DuplicateColumnMerger` - 按指定列合并重复行功能

**辅助工具类**:
- `com.tools.ExcelMerger` - 将多个 Excel 文件合并为一个文件的工具
- `com.oldFile.*` - 旧版实现 (CustomerIdMerger, ExcelDataFillingTool, TableMappingTool, ExcelProcessor)

### Core Features (GUI 五个标签页)

1. **Excel 处理器标签页** - 主要的 Excel 处理功能
   - 选择源数据文件和模板文件
   - 建立表头映射关系（支持多对一映射）
   - 数据转换和合并
   - 支持流式处理大文件

2. **建立映射关系标签页** - 构建和管理字段映射关系
   - 上传包含映射关系的 Excel 文件
   - 可视化配置源字段到目标字段的映射
   - 保存映射关系供后续使用
   - 支持大量映射关系的分块存储

3. **应用映射标签页** - 将已保存的映射关系应用到新数据
   - 选择查找字段和填充字段
   - 批量应用映射转换
   - 支持百万级数据的流式处理
   - 实时进度反馈

4. **跨表主键填充** - 根据主键在不同 Excel 表之间填充数据
   - A 表和 B 表主键逐行比对
   - 将 A 表的值字段填充到 B 表
   - 支持自定义主键列和值列

5. **重复列合并** - 按指定列合并重复行
   - 选择作为重复判断的关键列
   - 相同关键列的行合并为一行
   - 其他列内容用竖线 (|) 分隔拼接

### Technology Stack
- **GUI Framework**: Java Swing，包含现代化样式和进度对话框
- **Excel Processing**: Apache POI (5.2.3)，使用 SXSSF 支持大文件处理
- **Logging**: Log4j2 (2.20.0)
- **Build Tool**: Maven，使用 Shade plugin 创建 fat JAR
- **Packaging**: Launch4j 用于创建 Windows 可执行文件 (.exe)
- **Java Version**: Java 8 (兼容更高版本)

### Memory Management (内存管理优化)

**动态配置** - 系统根据可用内存自动调整参数：
- **低内存 (<512MB)**: Window 5 行, Batch 50 行, 最大文件 50MB
- **中等内存 (512MB-1GB)**: Window 20 行, Batch 200 行, 最大文件 200MB
- **高内存 (>1GB)**: Window 50 行, Batch 500 行, 最大文件 500MB

**关键优化特性**:
- 使用 SXSSF (Streaming Workbook) 配合动态窗口大小
- 自动内存监控，达到 85% 阈值时触发 GC
- 映射关系分块存储在临时文件中（不全部加载到内存）
- Batch 大小根据文件大小动态调整
- 低内存系统自动启用临时文件压缩

**处理能力**:
- 支持 .xlsx 和 .xls 格式
- 在低配置机器上处理 10 万+ 行数据
- 全程进度追踪和用户反馈

### Key Design Patterns (关键设计模式)
- **Multi-threading**: 使用 SwingWorker 进行后台处理并更新进度
- **Event-driven**: 完善的 GUI 事件处理，用于文件操作
- **Factory Pattern**: HeaderMapping 类支持不同的格式类型
- **Strategy Pattern**: 针对 .xlsx 和 .xls 文件的不同处理策略
- **Singleton**: MemoryConfigManager 用于集中管理内存配置
- **Streaming Pattern**: StreamingExcelProcessor 和 StreamingMappingApplier 使用 SAX 解析，避免大文件 OOM

## Development Notes (开发注意事项)

### File Processing Approach (文件处理方式)
- 使用 SAX 解析读取表头，最小化内存使用
- 同时支持 .xlsx 和 .xls 格式
- 对大数据集实施流式处理
- 全程进度追踪和用户反馈
- **分块处理**映射关系以防止内存溢出
- **StreamingExcelProcessor** 用于处理超大源文件（使用 SAX 读取）
- **StreamingMappingApplier** 用于应用映射到百万级数据

### GUI Architecture (GUI 架构)
- 标签页界面，不同功能分离
- 实时进度指示器和日志记录
- 文件验证和用户友好的错误处理
- 日志中显示内存状态
- 一致的样式和响应式布局

### Data Mapping System (数据映射系统)
- 灵活的表头映射，支持多对一关系
- 不同的合并格式（逗号分隔 vs 竖线分隔带标签）
- 可持久化的映射关系，可保存和重用
- 拖拽式选择的可视化映射界面
- **大量映射集存储在临时文件中以减少内存占用**

### Testing (测试)
- 当前项目中没有自动化单元测试
- 主要通过 GUI 进行功能测试和验证
- 建议在添加新功能时考虑添加测试用例

## Memory Optimization Implementation (内存优化实现)

### When to Use Optimized Features (何时使用优化功能)
- 处理 >50MB 的文件
- 在 <2GB RAM 的机器上工作
- 处理 >5 万行的数据
- 构建大量映射关系（>1 万条映射）

### Key Components (关键组件)

**MemoryConfigManager**:
- 自动检测系统内存并设置最优参数
- 监控实时内存使用情况
- 达到 85% 阈值时触发 GC
- 提供内存状态日志记录

**ChunkedMappingProcessor**:
- 将大量映射关系分块处理
- 将块存储在临时文件中（程序退出时自动清理）
- 处理映射时不将全部加载到内存
- 适用于 >1 万条映射关系

**StreamingExcelProcessor**:
- 使用 SAX 模式读取，SXSSF 模式写入
- 逐行处理，不将整个文件加载到内存
- 适用于超大源文件处理

**StreamingMappingApplier**:
- 流式应用映射关系
- 支持百万级数据处理
- 批量处理配合内存监控

### Troubleshooting OOM Issues (解决 OOM 问题)
1. 使用 `start_optimized.bat` 进行自动内存配置（**推荐**）
2. 检查日志中的内存状态和警告信息
3. 关闭其他占用内存的应用程序
4. 考虑将超大文件（>500MB）拆分为较小的块
5. 确保安装了 64 位 Java 运行时以获得更好的内存管理

### Temporary Files (临时文件)
- 位置：系统临时目录
- 前缀：`excel_mapping_`
- 自动清理：程序退出时
- 手动清理：删除临时文件夹中以 `excel_mapping_` 开头的目录

## Important Implementation Details (重要实现细节)

### SAX Parsing for Headers (SAX 解析表头)
`ExcelTools` 中的 `HeaderHandler` 内部类使用 SAX 解析仅读取 Excel 文件的表头行，避免将整个文件加载到内存。

### SXSSF Configuration (SXSSF 配置)
动态窗口大小通过 `MemoryConfigManager.getMemoryWindowSize()` 设置，而不是硬编码值。低内存系统会自动启用压缩。

### Batch Processing (批处理)
批处理大小根据系统内存和文件大小动态计算，使用 `MemoryConfigManager.getOptimalBatchSize(fileSizeMB)`。

### Multi-Format Merging (多格式合并)
当多个源表头映射到同一个目标表头时，`HeaderMapping.FormatType` 决定分隔符：
- `COMMA_SEPARATED`: 用于"手机号码"等字段（使用逗号分隔）
- `PIPE_SEPARATED`: 默认格式，包含源表头标签（格式：`源表头:数据 | 源表头:数据`）

### Streaming Components (流式组件)
- `StreamingExcelProcessor`: 用于 Excel 处理器标签页的大文件处理
- `StreamingMappingApplier`: 用于应用映射标签页的百万级数据处理
- 两者都使用 SAX 事件驱动解析，配合 SXSSF 写入

## Common Issues and Solutions (常见问题和解决方案)

**问题**: 处理时出现 OOM 错误
- **解决方案**: 使用 `start_optimized.bat` 或设置适当的 JVM 参数

**问题**: 加载表头失败 - "Tried to allocate an array" 错误
- **原因**: Excel文件的共享字符串表（SharedStringsTable）过大，超出POI的默认内存限制（1GB）
- **解决方案**:
  1. **推荐**: 使用 `start_optimized.bat` 启动，该脚本会根据系统内存自动分配合适的堆内存（最高4GB）
  2. **手动**: 使用命令 `java -Xmx4g -jar target\excel-processor-1.0-SNAPSHOT.jar` 手动指定更大内存
  3. **临时**: 如果文件特别大，尝试将Excel文件拆分为多个较小的文件
  4. **注意**: 程序已自动将POI数组限制提高到可用内存的70%或2GB（取较大值）
- **技术细节**: POI在解析XLSX文件时需要加载整个共享字符串表到内存，包含大量唯一字符串的文件会导致分配大数组

**问题**: 处理速度慢
- **解决方案**: 检查磁盘空间（需要临时文件），使用优化启动器

**问题**: 映射关系占用过多内存
- **解决方案**: 对于大量映射集（>1 万条）会自动使用 `ChunkedMappingProcessor`

**问题**: 文件过大错误
- **解决方案**: 错误信息会显示基于当前内存的最大文件大小；增加 JVM 堆内存或拆分文件

**问题**: 需要合并多个 Excel 文件
- **解决方案**: 使用 `com.tools.ExcelMerger` 工具类（需要修改硬编码路径）

## Additional Utilities (附加工具)

### ExcelMerger (Excel 合并器)
- 位置: `com.tools.ExcelMerger`
- 功能: 将一个文件夹中的多个 Excel 文件合并为单个文件
- 使用方式: 修改 `main` 方法中的 `folderPath`，然后运行
- 输出: 在同一文件夹中生成 `combined_output.xlsx`
- 注意: 当前版本使用硬编码路径，需要根据实际情况修改