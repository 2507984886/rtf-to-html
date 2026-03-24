# rtf-to-html

高质量 RTF 转 HTML 转换器，针对 Outlook 邮件优化，纯 C++17 实现，无第三方依赖。

## 目标

替代 Aspose（闭源）和 LibreOffice（体积过大），提供轻量级、高质量的 RTF→HTML 转换，支持在客户端集成（Linux、Android、HarmonyOS、iOS）。

## 功能特性

### 第一阶段（当前）

- **Outlook fromhtml 优化路径**：自动检测 `\fromhtml1`，从 `\*\htmltag` 组还原原始 HTML，近乎完美重建发件方的原始 HTML 邮件
- **完整 RTF 转换**：覆盖非 fromhtml 格式的 RTF 文档
  - 字符格式：粗体、斜体、下划线、删除线、上标/下标、颜色、高亮、字体、字号、字符间距
  - 段落格式：对齐方式、缩进、行间距、段前/段后间距
  - 表格：完整的行/列/单元格支持，背景色、边框、垂直合并
  - 图片：PNG/JPEG 转 base64 data URI 内嵌到 HTML
  - 超链接：`\field` / `\fldinst HYPERLINK` 提取
  - 列表：有序/无序列表嵌套
- **多代码页支持**：CP1252/1250/1251/1253/1254/1255/1256/1257/1258/850/437
- **Unicode 支持**：`\uN` 控制字和 `\'XX` 十六进制转义
- **跨平台**：仅使用 C++17 标准库，无平台特定 API

## 构建

```bash
cmake -B build -DCMAKE_BUILD_TYPE=Release
cmake --build build
```

## 使用

### API

```cpp
#include <rtf2html/rtf2html.h>

rtf2html::ConvertOptions opts;
opts.generateFullPage = true;  // 生成完整 HTML 页面
opts.embedImages = true;       // 内嵌图片（base64）

auto result = rtf2html::convert(rtfData.data(), rtfData.size(), opts);
if (result.success) {
    // 使用 result.html
}
```

### CLI 工具

```bash
# 转换文件
./rtf2html input.rtf -o output.html

# 生成完整 HTML 页面
./rtf2html -f input.rtf -o output.html

# 从 stdin 读取
cat input.rtf | ./rtf2html -f -o output.html
```

## 架构

```
include/rtf2html/rtf2html.h   - 公共 API
src/tokenizer.h/cpp           - RTF 词法分析器
src/state.h                   - RTF 状态结构（格式状态、目标枚举）
src/codepage.h/cpp            - 代码页 → UTF-8 转换表
src/converter.h/cpp           - 核心转换引擎（状态机 + HTML 生成）
src/rtf2html.cpp              - 公共 API 实现
tools/rtf2html_cli/main.cpp   - 命令行工具
```

## 开发路线

- **第一阶段**：按微软 RTF 规范编写核心转换引擎（✅ 完成）
- **第二阶段**：提供 Outlook 真实邮件样例，对比 Aspose 输出，针对性优化
