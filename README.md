# Excel工具集

<div align="center">

![Java](https://img.shields.io/badge/Java-8%2B-orange)
![Apache POI](https://img.shields.io/badge/Apache%20POI-5.2.3-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20MacOS-lightgrey)

**功能强大的Excel批量处理工具**

支持大文件处理 | 智能内存优化 | 流式处理 | 测试数据生成

[功能特性](#功能特性) • [快速开始](#快速开始) • [使用指南](#使用指南) • [常见问题](#常见问题)

</div>

---

## 📖 目录

- [项目简介](#项目简介)
- [功能特性](#功能特性)
- [技术亮点](#技术亮点)
- [系统要求](#系统要求)
- [快速开始](#快速开始)
- [使用指南](#使用指南)
- [启动方式](#启动方式)
- [测试数据生成器](#测试数据生成器)
- [内存优化](#内存优化)
- [项目结构](#项目结构)
- [常见问题](#常见问题)
- [更新日志](#更新日志)
- [贡献指南](#贡献指南)
- [许可证](#许可证)

---

## 项目简介

Excel工具集是一个基于Java Swing开发的**专业级Excel批量处理工具**，专注于解决企业级数据处理场景中的复杂需求。通过先进的流式处理技术和智能内存管理，能够**轻松处理百万级数据**和**数百MB的大文件**。

### 为什么选择Excel工具集？

- ✅ **处理大文件无压力** - 支持250MB+超大Excel文件，采用SAX流式解析
- ✅ **智能内存管理** - 自动检测系统内存，动态分配最优JVM参数
- ✅ **功能全面** - 5大核心功能模块，覆盖常见Excel处理场景
- ✅ **高性能** - SXSSF流式写入，处理百万行数据不卡顿
- ✅ **易于使用** - 图形化界面，拖拽操作，无需编程基础
- ✅ **开箱即用** - 提供测试数据生成器，一键生成百万级测试数据

---

## 功能特性

### 🎯 核心功能（5大模块）

#### 1. Excel处理器
**表头映射与数据转换的核心工具**

- ✨ 支持源数据文件和模板文件的表头映射
- ✨ 支持多对一映射（多个源表头→一个目标表头）
- ✨ 自动合并数据，保留源表头信息
- ✨ 支持大文件流式处理
- ✨ 实时显示处理进度

**使用场景**：
- 将多个不同格式的数据源统一转换为标准格式
- 整合多个系统导出的Excel数据
- 数据清洗和标准化

#### 2. 建立映射关系
**可视化配置字段映射关系**

- 📋 上传包含映射关系的Excel文件
- 📋 可视化配置源字段→目标字段映射
- 📋 支持保存映射关系供后续使用
- 📋 支持大量映射关系分块存储（>1万条）

**使用场景**：
- 建立字段转换规则库
- 配置数据字典映射
- 代码值与描述值对应关系

#### 3. 应用映射
**批量应用已保存的映射关系**

- 🔄 选择查找字段和填充字段
- 🔄 批量应用映射转换
- 🔄 支持百万级数据流式处理
- 🔄 实时进度反馈

**使用场景**：
- 批量翻译代码为中文描述
- 数据标准化和规范化
- 字段值替换和转换

#### 4. 跨表主键填充
**基于主键在不同Excel表之间填充数据**

- 🔗 A表和B表主键逐行比对
- 🔗 将A表的值字段填充到B表
- 🔗 支持自定义主键列和值列
- 🔗 高效的查找算法

**使用场景**：
- 订单表与客户表关联
- 主数据补充
- 多源数据合并

#### 5. 重复列合并
**按指定列合并重复行**

- 🔀 选择关键列作为重复判断依据
- 🔀 相同关键列的行合并为一行
- 🔀 其他列内容用竖线(|)分隔拼接
- 🔀 去重与聚合一步完成

**使用场景**：
- 客户信息去重
- 订单明细汇总
- 多次联系记录合并

---

## 技术亮点

### 🚀 性能优化

- **SAX流式解析** - 读取Excel表头，避免加载整个文件到内存
- **SXSSF流式写入** - 内存窗口动态调整，支持百万行数据写入
- **分块处理** - 大量映射关系分块存储在临时文件，避免OOM
- **批量处理** - 批处理大小根据文件大小动态计算
- **G1垃圾回收** - 针对大堆内存优化的GC策略

### 🧠 智能内存管理

- **动态内存检测** - 自动检测系统可用内存
- **分档配置** - 根据系统内存自动分配JVM参数（512MB-8GB）
- **POI数组限制提升** - 自动将Apache POI数组限制提升至可用内存的70%
- **内存监控** - 实时监控内存使用，达到85%阈值自动触发GC

### 💾 大文件支持

- **文件大小** - 支持处理250MB+的超大Excel文件
- **数据行数** - 支持处理100万+行数据
- **SharedStringsTable优化** - 特殊处理Excel字符串表，避免内存溢出
- **错误恢复** - 完善的异常处理和用户提示

---

## 系统要求

### 最低配置
- **Java版本**: Java 8 或更高版本（推荐Java 11+）
- **内存**: 2GB RAM
- **磁盘空间**: 100MB（不含数据文件）
- **操作系统**: Windows 7+ / Linux / MacOS

### 推荐配置
- **Java版本**: Java 11+ (64位)
- **内存**: 8GB+ RAM
- **磁盘空间**: 1GB+（处理大文件需要临时空间）
- **操作系统**: Windows 10+ / Linux / MacOS

### 大文件处理配置
- **内存**: 8GB+ RAM
- **Java**: 64位JDK
- **文件系统**: SSD（提升IO性能）

---

## 快速开始

### 1. 克隆项目

```bash
git clone <repository-url>
cd tool
```

### 2. 构建项目

```bash
# 编译
mvn clean compile

# 打包
mvn clean package
```

### 3. 启动应用

**方式1: 智能启动器（推荐）**
```bash
启动器.bat
```

**方式2: 自动优化启动**
```bash
start_optimized.bat
```

**方式3: 手动启动**
```bash
java -jar target/excel-processor-1.0-SNAPSHOT.jar
```

### 4. 快速体验

1. 启动应用后会看到5个标签页界面
2. 点击"Excel处理器"标签页
3. 选择示例文件开始处理

---

## 使用指南

### Excel处理器使用流程

#### Step 1: 选择文件
```
1. 点击"选择源数据文件"按钮
2. 选择包含原始数据的Excel文件
3. 点击"选择模板文件"按钮
4. 选择目标格式的模板Excel文件
```

#### Step 2: 加载表头
```
1. 点击"加载表头"按钮
2. 程序使用SAX方式快速读取表头
3. 左侧显示源数据表头
4. 右侧显示模板表头
```

#### Step 3: 配置映射
```
1. 在左侧列表选择一个或多个源表头
2. 在右侧列表选择一个目标表头
3. 点击"添加映射"按钮
4. 映射关系显示在下方列表
5. 可添加多个映射关系
```

#### Step 4: 开始处理
```
1. 点击"开始处理"按钮
2. 选择输出文件保存位置
3. 查看进度条和日志
4. 等待处理完成
```

### 建立映射关系使用流程

#### Step 1: 上传映射表
```
1. 准备包含映射关系的Excel文件
   格式：第一列为源字段，其他列为目标字段
2. 点击"上传关系表"按钮
3. 选择映射关系Excel文件
```

#### Step 2: 配置映射
```
1. 左侧选择源字段列
2. 右侧选择目标字段列
3. 点击"保存映射关系"
```

#### Step 3: 查看结果
```
1. 下方列表显示已保存的映射关系
2. 映射关系自动持久化
3. 支持大量映射（>1万条）自动分块存储
```

### 应用映射使用流程

#### Step 1: 上传目标表
```
1. 点击"上传目标表"按钮
2. 选择需要应用映射的Excel文件
```

#### Step 2: 选择字段
```
1. 选择查找字段（用于匹配映射关系的字段）
2. 选择填充字段（用于写入映射结果的字段）
```

#### Step 3: 应用映射
```
1. 点击"应用映射"按钮
2. 选择输出文件位置
3. 程序使用流式处理应用映射
4. 支持百万级数据
```

---

## 启动方式

### 🎯 三种启动方式对比

| 启动方式 | 内存分配 | 适用场景 | 推荐度 |
|---------|---------|---------|--------|
| **启动器.bat** | 用户选择 | 所有场景 | ⭐⭐⭐⭐⭐ |
| **start_optimized.bat** | 自动检测 | 自动适配 | ⭐⭐⭐⭐ |
| **双exe方案** | 固定配置 | 分发给用户 | ⭐⭐⭐⭐ |

### 方式1: 智能启动器（推荐）

**文件**: `启动器.bat`

**特点**:
- 图形化菜单选择
- 4种内存模式
- 用户友好

**使用方法**:
```bash
启动器.bat
```

**菜单选项**:
```
[1] 快速启动（<100MB文件）    - 默认内存
[2] 标准启动（100-200MB文件） - 智能检测
[3] 高性能启动（200MB+文件）  - 6GB内存
[4] 极致性能启动（500MB+文件）- 8GB内存
```

### 方式2: 自动优化启动

**文件**: `start_optimized.bat`

**特点**:
- 自动检测系统内存
- 智能分配JVM参数
- 无需手动选择

**内存分配规则**:
```
系统内存 <1GB   → 分配 512MB
系统内存 1-2GB  → 分配 1GB
系统内存 2-4GB  → 分配 2.5GB
系统内存 4-8GB  → 分配 6GB
系统内存 >8GB   → 分配 8GB
```

### 方式3: 打包exe方案

**生成标准版exe**:
```bash
launch4jc launch4j-config.xml
# 生成：ExcelProcessor.exe (4GB内存)
```

**生成大文件版exe**:
```bash
launch4jc launch4j-config-largefile.xml
# 生成：ExcelProcessor_LargeFile.exe (6GB内存)
```

**使用说明**:
```
文件 <200MB  → 使用 ExcelProcessor.exe
文件 ≥200MB  → 使用 ExcelProcessor_LargeFile.exe
```

---

## 测试数据生成器

### 🎲 快速生成测试数据

**功能**: 生成真实的百万级Excel测试数据，用于测试工具性能。

**特点**:
- ✅ 真实中文姓名（100个姓氏 + 200个名字）
- ✅ 真实手机号（符合中国号段规则）
- ✅ 30种客户备注模板
- ✅ 详细的跟进记录（3-5条历史记录）
- ✅ 可自定义文件大小

### 使用方法

**默认生成150MB文件**:
```bash
java -cp target/excel-processor-1.0-SNAPSHOT.jar com.tools.TestDataGenerator
```

**指定文件大小**:
```bash
# 生成10MB测试文件
java -cp target/excel-processor-1.0-SNAPSHOT.jar com.tools.TestDataGenerator 10

# 生成250MB测试文件
java -cp target/excel-processor-1.0-SNAPSHOT.jar com.tools.TestDataGenerator 250
```

### 生成的数据格式

| 字段名 | 描述 | 示例 |
|--------|------|------|
| 客户姓名 | 真实中文姓名 | 张伟、李娜、王芳 |
| 客户电话 | 11位手机号 | 13912345678 |
| 客户备注 | 30种业务模板 | "意向客户，待进一步跟进" |
| 客户星级 | 1-5星评级 | ★★★★☆ 或 4星 |
| 其他备注 | 150-250字详细备注 | 客户背景、需求、预算信息 |
| 跟进记录 | 3-5条历史记录 | 日期、操作类型、详细内容 |

### 性能指标

- **生成速度**: 15,000-25,000行/分钟
- **内存占用**: <100MB（使用SXSSF流式写入）
- **150MB文件**: 约780,000行，耗时3-5分钟
- **输出位置**: 项目根目录，文件名带时间戳

---

## 内存优化

### 🧠 智能内存管理机制

#### 动态配置

**MemoryConfigManager** - 核心内存管理器

```java
// 自动检测系统内存
MemoryConfigManager.getInstance().getMemoryStatus();

// 获取动态配置参数
int windowSize = memoryConfig.getMemoryWindowSize();
int batchSize = memoryConfig.getBatchSize();
int maxFileSize = memoryConfig.getMaxFileSize();
```

**配置规则**:
```
低内存 (<512MB):  Window 5行,  Batch 50行,  最大文件 50MB
中等内存 (512MB-1GB): Window 20行, Batch 200行, 最大文件 200MB
高内存 (>1GB):    Window 50行, Batch 500行, 最大文件 500MB
```

#### 关键优化技术

1. **SXSSF流式写入**
   - 内存窗口动态调整
   - 只在内存中保留N行数据
   - 自动刷写到磁盘

2. **SAX流式读取**
   - 事件驱动解析
   - 不加载整个文件到内存
   - 只读取必要的表头信息

3. **分块处理**
   - 大量映射关系分块存储
   - 临时文件存储中间结果
   - 程序退出自动清理

4. **内存监控**
   - 实时监控内存使用率
   - 达到85%阈值触发GC
   - 防止OOM异常

#### POI数组限制优化

```java
// 自动提升POI数组限制
long availableMemory = Runtime.getRuntime().maxMemory();
long maxArraySize = Math.min((long)(availableMemory * 0.7), Integer.MAX_VALUE - 8);
IOUtils.setByteArrayMaxOverride((int) maxArraySize);
```

**效果**:
- 默认限制：1GB
- 优化后：可用内存的70% 或 2GB（取较大值）
- 最大支持：约2.1GB（Integer.MAX_VALUE限制）

### 处理能力

| 系统配置 | JVM内存 | 支持文件大小 | 支持行数 |
|---------|---------|-------------|---------|
| 2GB RAM | 1GB | ~100MB | ~50万行 |
| 4GB RAM | 2.5GB | ~150MB | ~80万行 |
| 8GB RAM | 6GB | ~250MB | ~130万行 |
| 16GB RAM | 8GB | ~400MB | ~200万行 |

---

## 项目结构

```
tool/
├── src/
│   ├── main/
│   │   ├── java/
│   │   │   ├── com.xzy/                    # 核心包
│   │   │   │   ├── ExcelTools.java         # 主程序入口
│   │   │   │   ├── MemoryConfigManager.java     # 内存配置管理器
│   │   │   │   ├── ChunkedMappingProcessor.java # 分块映射处理器
│   │   │   │   ├── StreamingExcelProcessor.java # 流式Excel处理器
│   │   │   │   ├── StreamingMappingApplier.java # 流式映射应用器
│   │   │   │   ├── CrossSheetFiller.java        # 跨表填充工具
│   │   │   │   └── DuplicateColumnMerger.java   # 重复列合并工具
│   │   │   ├── com.tools/                  # 工具包
│   │   │   │   ├── TestDataGenerator.java  # 测试数据生成器
│   │   │   │   └── ExcelMerger.java        # Excel合并工具
│   │   │   └── com.oldFile/                # 旧版实现
│   │   └── resources/
│   │       └── META-INF/
│   │           └── MANIFEST.MF
│   └── log4j2.xml                          # 日志配置
├── target/                                  # 编译输出
│   └── excel-processor-1.0-SNAPSHOT.jar
├── 启动器.bat                               # 智能启动器
├── start_optimized.bat                      # 自动优化启动
├── launch4j-config.xml                      # Launch4j标准版配置
├── launch4j-config-largefile.xml            # Launch4j大文件版配置
├── pom.xml                                  # Maven配置
├── README.md                                # 本文档
├── CLAUDE.md                                # 开发者文档
├── AGENTS.md                                # 仓库规范
├── 启动方式指南.md                          # 启动方式详细说明
├── TestDataGenerator使用说明.md            # 测试数据生成器说明
└── .gitignore
```

### 核心类说明

| 类名 | 职责 | 关键特性 |
|-----|------|---------|
| `ExcelTools` | 主程序入口，GUI管理 | Swing界面，5个标签页 |
| `MemoryConfigManager` | 内存配置管理 | 单例，动态检测系统内存 |
| `ChunkedMappingProcessor` | 分块映射处理 | 避免大量映射OOM |
| `StreamingExcelProcessor` | 流式Excel处理 | SAX读取 + SXSSF写入 |
| `StreamingMappingApplier` | 流式映射应用 | 支持百万级数据 |
| `CrossSheetFiller` | 跨表填充 | 主键匹配填充 |
| `DuplicateColumnMerger` | 重复行合并 | 按列去重聚合 |
| `TestDataGenerator` | 测试数据生成 | 真实数据，流式生成 |

---

## 常见问题

### Q1: 处理大文件时出现OOM错误

**问题**: `OutOfMemoryError` 或 `java.lang.OutOfMemoryError: Java heap space`

**解决方案**:
```bash
# 方案1: 使用优化启动脚本（推荐）
start_optimized.bat

# 方案2: 手动指定更大内存
java -Xmx8g -jar target/excel-processor-1.0-SNAPSHOT.jar

# 方案3: 使用启动器选择高性能模式
启动器.bat
→ 选择 [3] 或 [4]
```

### Q2: 加载表头失败 - "Tried to allocate an array" 错误

**原因**: Excel文件的SharedStringsTable过大，超出POI默认1GB限制

**解决方案**:
```bash
# 程序已自动提升限制到可用内存的70%或2GB
# 如果仍然失败，使用start_optimized.bat启动
start_optimized.bat

# 或手动指定大内存
java -Xmx6g -jar target/excel-processor-1.0-SNAPSHOT.jar
```

**技术细节**:
- POI解析XLSX时需加载整个SharedStringsTable到内存
- 包含大量唯一字符串的文件会导致该表很大
- 程序已通过 `IOUtils.setByteArrayMaxOverride()` 提升限制

### Q3: 处理速度慢

**可能原因**:
1. 磁盘IO慢（使用机械硬盘）
2. 内存不足导致频繁GC
3. 文件过大

**优化建议**:
```bash
# 1. 使用SSD存储临时文件
# 2. 关闭其他占用内存的程序
# 3. 使用优化启动脚本分配更多内存
start_optimized.bat

# 4. 将超大文件拆分处理
```

### Q4: 启动器选择哪个模式？

**推荐规则**:
```
文件大小 <100MB    → [1] 快速启动
文件大小 100-200MB → [2] 标准启动（推荐）
文件大小 200-300MB → [3] 高性能启动
文件大小 >300MB    → [4] 极致性能启动
```

### Q5: 如何生成测试数据？

```bash
# 生成10MB测试文件（快速测试）
java -cp target/excel-processor-1.0-SNAPSHOT.jar com.tools.TestDataGenerator 10

# 生成150MB测试文件（性能测试）
java -cp target/excel-processor-1.0-SNAPSHOT.jar com.tools.TestDataGenerator 150

# 生成250MB测试文件（压力测试）
java -cp target/excel-processor-1.0-SNAPSHOT.jar com.tools.TestDataGenerator 250
```

### Q6: exe启动失败

**问题**: 双击exe没反应或报错

**检查清单**:
1. ✅ 是否安装了Java 8+
2. ✅ 是否是64位Java（处理大文件需要）
3. ✅ target目录下是否有jar文件

**解决方案**:
```bash
# 改用bat启动查看详细错误
start_optimized.bat

# 或查看是否是内存不足
# 使用大文件版exe（如果已生成）
ExcelProcessor_LargeFile.exe
```

### Q7: 映射关系太多导致内存不足

**问题**: 保存超过1万条映射关系时OOM

**解决方案**:
- ✅ 程序已自动使用 `ChunkedMappingProcessor`
- ✅ 映射关系自动分块存储在临时文件
- ✅ 不会全部加载到内存

**技术细节**:
```java
// 程序会自动检测映射关系数量
if (mappingCount > 10000) {
    // 自动启用分块处理器
    chunkedProcessor.processMappings();
}
```

### Q8: 临时文件在哪里？

**位置**: 系统临时目录
```
Windows: C:\Users\<用户名>\AppData\Local\Temp\excel_mapping_*
Linux:   /tmp/excel_mapping_*
Mac:     /var/folders/.../excel_mapping_*
```

**清理**: 程序退出时自动清理

### Q9: 如何处理500MB+超大文件？

**建议方案**:
```bash
# 1. 使用极致性能模式
启动器.bat → 选择 [4]

# 2. 或手动指定10-12GB内存（如果系统有16GB+）
java -Xms2g -Xmx12g -XX:+UseG1GC -jar target/excel-processor-1.0-SNAPSHOT.jar

# 3. 或考虑将文件拆分为多个小文件
# 使用Excel或Python脚本拆分
```

### Q10: 支持哪些Excel格式？

**支持格式**:
- ✅ `.xlsx` (推荐，性能更好)
- ✅ `.xls` (旧格式，兼容)

**不支持**:
- ❌ `.xlsm` (带宏)
- ❌ `.xlsb` (二进制)
- ❌ `.csv` (纯文本)

---

## 更新日志

### v1.0.0 (2025-10-19)

#### 🎉 核心功能
- ✅ Excel处理器 - 表头映射与数据转换
- ✅ 建立映射关系 - 可视化字段映射配置
- ✅ 应用映射 - 批量应用映射关系
- ✅ 跨表主键填充 - 基于主键填充数据
- ✅ 重复列合并 - 按列去重聚合

#### 🚀 性能优化
- ✅ 内存优化 - 智能检测系统内存，动态分配JVM参数
- ✅ 流式处理 - SAX读取 + SXSSF写入，支持百万级数据
- ✅ 分块处理 - 大量映射关系分块存储，避免OOM
- ✅ POI优化 - 自动提升数组限制到可用内存的70%

#### 🛠️ 工具增强
- ✅ 测试数据生成器 - 一键生成百万级真实测试数据
- ✅ 智能启动器 - 图形化菜单，4种内存模式选择
- ✅ 自动优化启动 - 自动检测系统内存并分配最优配置
- ✅ 双exe方案 - 标准版 + 大文件版，适合不同场景

#### 📖 文档完善
- ✅ README.md - 完整的项目文档
- ✅ CLAUDE.md - 开发者指南和技术细节
- ✅ 启动方式指南.md - 启动方式详细说明
- ✅ TestDataGenerator使用说明.md - 测试数据生成器文档
- ✅ AGENTS.md - 仓库规范和编码标准

#### 🐛 问题修复
- ✅ 修复 SharedStringsTable 过大导致的OOM
- ✅ 修复 Launch4j 配置中的JAR文件名错误
- ✅ 修复 start_optimized.bat 内存档位不足问题

---

## 贡献指南

欢迎贡献代码、报告问题或提出建议！

### 如何贡献

1. **Fork本仓库**
2. **创建特性分支** (`git checkout -b feature/AmazingFeature`)
3. **提交更改** (`git commit -m 'feat: Add some AmazingFeature'`)
4. **推送到分支** (`git push origin feature/AmazingFeature`)
5. **打开Pull Request**

### 提交规范

```
feat: 新功能
fix: 修复bug
perf: 性能优化
refactor: 重构
docs: 文档更新
test: 测试相关
chore: 构建/工具相关
```

### 开发环境

```bash
# 克隆仓库
git clone <repository-url>

# 安装依赖
mvn clean install

# 运行测试
mvn test

# 启动开发版本
mvn exec:java -Dexec.mainClass="com.xzy.ExcelTools"
```

### 代码规范

- ✅ Java 8语法
- ✅ 4空格缩进
- ✅ UTF-8编码
- ✅ 详细的注释
- ✅ 遵循阿里巴巴Java开发手册

详见 [AGENTS.md](AGENTS.md)

---

## 许可证

本项目采用 **MIT 许可证**。详见 [LICENSE](LICENSE) 文件。

---

## 联系方式

- 📧 Email: [2894702202@qq.com]
- 🐛 Issues: [GitHub Issues]
- 📖 WX: [xzynet-top]

---

## 致谢

- [Apache POI](https://poi.apache.org/) - Excel文件处理
- [Log4j2](https://logging.apache.org/log4j/2.x/) - 日志框架
- [Launch4j](http://launch4j.sourceforge.net/) - Java打包工具

---

<div align="center">

**如果这个项目对您有帮助，请给一个⭐️Star⭐️支持一下！**

Made with ❤️ by  XZY & Course & Claude & Codex

</div>
