# Repository Guidelines

## 项目结构与模块组织
- 源码: `src/main/java` (包 `com.xzy`, `com.tools`, `com.oldFile`).
- 资源: `src/main/resources` (含 `META-INF/MANIFEST.MF`); 日志配置位于 `src/log4j2.xml`.
- 构建产物: `target/` (编译类与 shaded JAR).
- 入口类: `com.xzy.ExcelTools` (由 Maven Shade 指定).

## 构建、测试与本地开发命令
- `mvn clean package` - 构建可运行的单 JAR 至 `target/`.
- `mvn compile` - 编译源码至 `target/classes`.
- `java -jar target/excel-processor-1.0-SNAPSHOT-shaded.jar` - 启动 GUI 工具.
- Windows: `start_optimized.bat` - 以内存优化参数启动.
- 可选: 在 `pom.xml` 保持 `project.build.sourceEncoding=UTF-8`.

## 编码风格与命名约定
- Java 8; 4 空格缩进; 禁止 Tab.
- 类名 PascalCase; 方法/字段 camelCase; 常量 UPPER_SNAKE_CASE.
- 包路径保持在 `com.xzy`, `com.tools`, `com.oldFile`.
- 使用 Log4j2 记录日志; 避免在核心逻辑中使用 `System.out.println`.
- 文件编码统一 UTF-8 (无 BOM); 标点使用半角; 若出现乱码请以 UTF-8 重新保存.

## 测试指南
- 建议使用 JUnit 5.
- 将测试放在 `src/test/java`, 与源码包结构对应.
- 测试文件命名为 `*Test.java`; 保持可重复、快速.
- 面向大文件功能, 优先使用小型合成数据与流式 API.

## 提交与合并请求规范
- 提交前缀: `feat:`, `fix:`, `perf:`, `refactor:`; 信息使用祈使句、清晰.
- PR 需包含描述、关联 Issue、复现步骤; GUI 变更附截图/GIF.
- 合并前确保 `mvn clean package` 通过且 JAR 可本地运行.

## 安全与配置提示
- 对不可信的 Excel 输入进行校验 (表头、行数、文件大小).
- 避免从表格加载外部资源; 非必要时禁用公式计算.
- 通过 `MemoryConfigManager` 记录内存配置修改.

## Agent 专用说明
- 回答语言: 使用中文, 技术相关名词保留英文原文.
- 编码要求: 全仓库统一 UTF-8 (无 BOM), 使用半角标点, 严禁产生中文乱码; 如出现乱码请以 UTF-8 重新保存文件.
- 文档约定: 本 AGENTS.md 及后续修改保持中文撰写; 其他文档亦优先中文说明 (技术名词保留英文原文).
