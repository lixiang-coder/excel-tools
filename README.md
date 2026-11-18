# Excel工具集

一个功能强大的Excel处理工具，支持表头映射、数据合并和批量处理等功能。

## 主要功能

### 1. Excel处理器
- 支持源数据文件和模板文件的表头映射
- 支持多个源表头映射到同一个模板表头
- 自动处理数据合并，并在合并时保留源表头信息
- 支持大文件处理，采用流式处理方式（SXSSF）
- 实时显示处理进度
- 支持批量数据处理

### 2. 建立映射关系

- 选择关系表的两列，建立值与值的对应关系。可以在本次运行中累计导入多张关系表。
- 相同的关系只计一次。可在列表中查看正向、反向冲突的各个候选值，以及它们来自哪个文件、哪一行。

### 3. 应用映射功能
- 支持将已配置的映射关系应用到新的Excel文件
- 支持批量处理大量数据
- 实时显示处理进度
- 自动处理数据格式转换

### 4. 跨表主键填充工具
- 支持A表与B表按主键逐行匹配
- 将源表数据列的值填充到目标表对应的列
- 支持大小写无关、自动去空格的模糊匹配
- 实时显示处理进度

### 5. 重复列合并工具
- 按指定列（如CustomerId）识别重复行
- 选中的合并列按首次出现顺序去重并用 `|` 拼接，其他列取首个非空值
- 支持自定义选择需要合并的列
- 适合批量去重、合并等场景

## 技术特点

- 使用Java Swing构建GUI界面
- 采用Apache POI 5.2.3处理Excel文件
- 支持大文件处理（使用SXSSF流式模式，100行内存窗口）
- 内存优化设计，支持处理百万级数据
- 所有耗时操作使用SwingWorker异步执行，不阻塞UI
- 详细的日志记录（Log4j2）

## 系统要求

### 客户电脑（使用安装包）

- Windows 10/11（64位）
- 不需要预装Java环境
- 内存：建议4GB以上
- 磁盘空间：安装包会内置精简Java运行时，建议预留300MB以上

### 开发/打包电脑

- Java 8或更高版本
- 生成免Java安装包需要JDK 17或更高版本（使用`jpackage`）
- 生成Windows `.exe`/`.msi`安装器需要WiX Toolset 3.x，或运行打包脚本时使用`-DownloadWix`
- Maven 3.x
- 操作系统：Windows/Linux/MacOS
- 内存：建议4GB以上
- 磁盘空间：根据处理文件大小而定

## 项目结构

```
src/
├── main/
│   ├── java/com/xzy/
│   │   ├── ExcelTools.java               ← 主程序（5 Tab GUI）
│   │   ├── CrossSheetFiller.java         ← 跨表填充（@Deprecated，被Tab调用）
│   │   └── DuplicateColumnMerger.java    ← 重复列合并（@Deprecated，被Tab调用）
│   └── resources/
│       ├── log4j2.xml                    ← 日志配置
│       └── META-INF/MANIFEST.MF
├── legacy/                               ← 旧版代码归档（不参与编译）
docs/                                     ← 项目文档
test-data/                                ← 测试数据
launch4j-config.xml                       ← Windows EXE打包配置
pom.xml                                   ← Maven构建配置
```

## 构建与运行

```bash
mvn clean compile          # 编译
mvn package                # 打包（生成 target/excel-processor-1.0-SNAPSHOT.jar）
java -jar target/excel-processor-1.0-SNAPSHOT.jar   # 运行
```

### 生成客户安装包（不要求客户安装Java）

推荐使用项目根目录的`build-installer.ps1`：

```powershell
# 生成 Windows .exe 安装器（默认，标准安装器，需要 WiX）
powershell -NoProfile -ExecutionPolicy Bypass -File .\build-installer.ps1 -DownloadWix

# 生成单文件自解压安装包（不需要 WiX，使用 Windows 内置 IExpress）
powershell -NoProfile -ExecutionPolicy Bypass -File .\build-installer.ps1 -Type self-extracting-exe

# 如果已安装 WiX Toolset，也可以指定 WiX bin 目录
powershell -NoProfile -ExecutionPolicy Bypass -File .\build-installer.ps1 -WixBin "C:\Program Files (x86)\WiX Toolset v3.14\bin"

# 只生成免安装目录，用于快速验证或内部分发
powershell -NoProfile -ExecutionPolicy Bypass -File .\build-installer.ps1 -Type app-image

# 指定中文安装包应用名
powershell -NoProfile -ExecutionPolicy Bypass -File .\build-installer.ps1 -AppName "Excel工具集" -DownloadWix
```

输出目录为`dist/`：

- `dist/ExcelTools-1.0.1.exe`：客户安装包
- `dist/ExcelTools-1.0.1-setup.exe`：自解压客户安装包（不依赖WiX）
- `dist/ExcelTools/ExcelTools.exe`：免安装应用目录（`-Type app-image`）

`AppVersion`会写入内置`ExcelTools.exe`的文件版本和产品版本，并用于安装包文件名。自解压安装包的外层由Windows IExpress生成，因此在文件属性中会显示系统`WEXTRACT.EXE`版本，而不是应用版本。

#### 自解压安装包的安装与卸载

`ExcelTools-1.0.1-setup.exe`使用当前Windows用户权限安装，不需要管理员权限。默认安装位置为：

```text
%LOCALAPPDATA%\Programs\ExcelTools
```

主程序及快捷方式位置如下：

```text
%LOCALAPPDATA%\Programs\ExcelTools\ExcelTools.exe
%USERPROFILE%\Desktop\ExcelTools.lnk
%APPDATA%\Microsoft\Windows\Start Menu\Programs\ExcelTools\ExcelTools.lnk
```

该自解压安装包不提供安装目录选择，也不会在安装完成后自动启动程序。重复安装时会先删除原安装目录，再写入新版本。

当前`self-extracting-exe`安装方式不会注册Windows“应用和功能”卸载项。卸载前先关闭ExcelTools，然后执行：

```powershell
Remove-Item "$env:LOCALAPPDATA\Programs\ExcelTools" -Recurse -Force
Remove-Item "$env:USERPROFILE\Desktop\ExcelTools.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\ExcelTools" -Recurse -Force -ErrorAction SilentlyContinue
```

卸载只删除程序目录和快捷方式，不会删除用户另行保存的Excel文件。

Launch4j仍可通过`launch4j-config.xml`将JAR打包为`ExcelProcessor.exe`，但该方式默认依赖客户电脑已有Java环境；客户零Java环境分发请优先使用`build-installer.ps1`。

## 使用说明

### Excel处理器

1. 选择文件
   - 点击"选择文件"按钮
   - 选择源数据文件（包含原始数据的Excel文件）
   - 选择模板文件（目标格式的Excel文件）

2. 配置映射关系
   - 在左侧列表中选择源表头
   - 在右侧列表中选择目标表头
   - 点击"添加映射"按钮建立映射关系
   - 可以建立多个映射关系

3. 保存文件
   - 点击"保存文件"按钮
   - 选择保存位置
   - 等待处理完成

### 建立映射关系

1. 上传关系表
   - 点击"上传关系表"按钮
   - 选择包含映射关系的Excel文件

2. 建立映射关系
   - 在左侧选择源字段
   - 在右侧选择目标字段
   - 点击"保存映射关系"按钮
   - 选中的关系字段含公式时，读取文件中已保存的计算结果。如果提示缺少结果，请按提示的文件、工作表和单元格重新计算并保存，再导入。本次导入失败时，之前导入的关系仍可使用。

3. 查看映射关系
   - 列表顶部显示冲突及来源。例如，同一编号对应两个姓名时，按编号查姓名会有冲突。两个编号对应同一姓名时，仍可按编号查姓名，按该姓名反查编号时会有冲突。
   - 可以继续导入其他关系表。“重置映射关系”会清空本次累计的关系和来源。

### 应用映射

1. 上传目标表
   - 点击"上传目标表"按钮
   - 选择需要应用映射的Excel文件

2. 选择字段
   - 选择查找字段（用于查找映射关系的字段）
   - 选择填充字段（用于填充映射值的字段）

3. 应用映射
   - 点击"应用映射"按钮
   - 选择保存位置
   - 程序先按正向查找。没有正向关系时再反查。匹配方向只有一个候选值时填入该值，两边都找不到时仍填入查找值。
   - 默认以目标文件的原格式另存一份，只填充首张表的选定列。原有工作表、公式、列宽、样式和合并区域随结果保留。填充格落在合并区域时，只有左上角可以写入。
   - 旧版 `.xls` 也可以将扩展名指定为 `.xlsx`。此转换支持普通表格及常用样式；遇到宏、图形、条件格式或数据验证等无法保留的内容时，程序会提示改存为 `.xls`。
   - 如果匹配方向有多个候选值，该行保留原值，其他行继续处理。冲突格含公式时，写入文件中已保存的计算结果。
   - 遇到冲突时，程序会在工作簿末尾追加“映射冲突”表，列出目标行号、查找值、匹配方向、候选值、关系文件和关系行号。已有同名表或明细写满时，自动使用下一个可用名称，如“映射冲突2”。

### 跨表主键填充

1. 选择源表和目标表
2. 配置匹配表头（源表和目标表的主键列）
3. 配置数据表头（源表数据列和目标表填充列）
4. 点击"开始填充"，选择输出路径

结果默认按目标文件的原格式另存。工具只在目标首张工作表追加匹配值，保留其他工作表和原有版式。填入格原本含公式时，会将已保存的计算结果与新增值拼成文本。用于匹配或拼接的公式缺少有效缓存时，按提示在 Excel 或 WPS 中重新计算并保存后重试。

### 重复列合并

1. 选择输入文件
2. 选择重复列（用于判断重复的列）
3. 勾选要合并的列
4. 点击"开始合并"，选择输出路径

重复列为空、但其他列有数据的行按输入的相对顺序逐条保留，不互相合并；完全空行会被忽略。重复列非空时仍按原规则合并。完成提示会显示有效输入行数、输出行数和保留的空键行数。

### 表格合并

自动识别遇到仅凭内容无法判定的首行，会提示选择合并方式。姓名、城市等无表头名单请选择“无表头合并（保留首行）”。首行确实是字段名时请选择“有表头合并”。按表头合并时，与标准表头没有共同列的工作表会在报告中标为失败。

按列名选择或对齐字段时，同一张表的字段名必须能区分。遇到重复名称，程序会指出冲突的两列，请先修改表头再重试。无表头合并按列的位置复制，可以保留重复的文字。

## 注意事项

1. 文件大小限制
   - 单个文件最大支持1000MB
   - 建议分批处理超大文件

2. 数据格式
   - 支持.xlsx和.xls格式
   - 建议使用.xlsx格式以获得更好的性能

3. 内存使用
   - 程序会自动优化内存使用
   - 处理大文件时建议关闭其他内存占用大的程序

4. 错误处理
   - 程序会自动记录错误日志
   - 遇到错误时会显示详细的错误信息

## 常见问题

1. 如何处理大文件？
   - “Excel 处理器”采用流式输出，并对源文件和模板分别执行 1000 MB 大小检查。
   - “应用映射”为保留原有工作表和版式，需要将整个工作簿载入内存。可处理的规模取决于工作簿内容与可用内存。

2. 如何处理多个表头映射到同一个目标表头？
   - 程序会自动合并数据
   - 合并时会保留源表头信息，格式为：源表头：数据 | 源表头：数据

3. 如何处理公式？
   - **表格合并**将公式在源工作簿中最后保存的计算结果写入总表。合并后移动行或调整列顺序，结果仍是源表的值。数字、文本、布尔值和 Excel 错误值保留各自的类型。
   - **Excel 处理器、应用映射**读取 `.xlsx` 时，以源文件保存的公式结果进行转换或匹配。若提示某个公式缺少计算结果，请在 Excel 或 WPS 中重新计算并保存源文件，再执行导出。如果刚改过公式引用的数据，也应先重算保存。

## 更新日志

### v1.0.1
- 优化Excel数值型日期的格式化逻辑，避免输出Java默认英文日期
- 日期时间统一输出为常用数字格式，并按实际精度保留分钟或秒
- 补充自解压安装包的安装目录和卸载说明

### v1.0.0
- 初始版本发布
- 支持基本的Excel处理功能
- 支持表头映射和数据合并
- 支持大文件处理

## 许可证

本项目采用MIT许可证。
