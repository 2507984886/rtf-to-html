#pragma once
// state.h - RTF 解析状态结构定义
// 包含字符状态、段落状态、目标类型等

#include <string>
#include <vector>
#include <cstdint>

namespace rtf2html {

// ============================================================
// 字体表项
// ============================================================
struct FontEntry {
    std::string name;       // 字体名称
    std::string altName;    // 备用名称
    int charset = 0;        // \fcharset 值
    int family = 0;         // \froman, \fswiss 等
    int codepage = 1252;    // 代码页
};

// ============================================================
// 颜色表项
// ============================================================
struct ColorEntry {
    uint8_t r = 0, g = 0, b = 0;
    bool isAuto = true;  // 第一个条目为自动颜色
};

// ============================================================
// 列表级别定义
// ============================================================
struct ListLevelDef {
    int numType = 0;         // 0=项目符号, 1=十进制, 2=小写字母 等
    int start = 1;           // 起始编号
    std::string bulletText;  // 项目符号文本
};

// ============================================================
// 列表定义
// ============================================================
struct ListDef {
    int listId = 0;
    std::vector<ListLevelDef> levels;
};

// ============================================================
// 表格单元格定义
// ============================================================
struct TableCellDef {
    int rightBoundary = 0;          // 右边界（twips，从左边距起）
    int bgColorIndex = -1;          // 背景颜色索引（-1=无）
    bool borderLeft = false;        // 左边框
    bool borderRight = false;       // 右边框
    bool borderTop = false;         // 上边框
    bool borderBottom = false;      // 下边框
    int borderLineWidth = 15;       // 边框宽度（twips）
    std::string borderColor;        // 边框颜色（十六进制）
    int valign = 0;                 // 垂直对齐: 0=顶, 1=居中, 2=底
    bool vertMergeFirst = false;    // 垂直合并第一单元格
    bool vertMerge = false;         // 垂直合并延续
};

// ============================================================
// 字符格式状态
// ============================================================
struct CharState {
    int fontIndex = 0;          // 字体索引
    int fontSize = 24;          // 字体大小（半磅）
    bool bold = false;          // 粗体
    bool italic = false;        // 斜体
    bool underline = false;     // 下划线
    bool strikethrough = false; // 删除线
    bool superscript = false;   // 上标
    bool subscript = false;     // 下标
    bool hidden = false;        // 隐藏文本
    bool allCaps = false;       // 全大写
    bool smallCaps = false;     // 小型大写字母
    int fgColorIndex = 0;       // 前景色索引（0=自动）
    int bgColorIndex = -1;      // 背景色索引（-1=无）
    int highlightIndex = -1;    // 高亮颜色索引
    int charSpacing = 0;        // 字符间距（twips）
    int charScaleX = 100;       // 水平缩放（百分比）
    int underlineStyle = 0;     // 下划线样式（0=无, 1=单线 等）
    int styleIndex = 0;         // 样式索引
    int unicodeSkip = 1;        // \uc 值
    int lang = 0;               // 语言
    int fontCharset = 0;        // 当前字体字符集
    int fontCodepage = 1252;    // 当前字体代码页
    // 特殊文字效果（Feature 14）
    bool outline = false;       // \outl 空心字
    bool shadow = false;        // \shad 阴影
    bool emboss = false;        // \embo 浮雕
    bool engrave = false;       // \impr 雕刻
    // 文字方向（Feature 9）
    bool rtl = false;           // \rtlch 从右到左字符运行
};

// ============================================================
// 段落边框（Feature 8）
// ============================================================
struct ParaBorder {
    bool left = false;          // 左边框
    bool right = false;         // 右边框
    bool top = false;           // 上边框
    bool bottom = false;        // 下边框
    bool box = false;           // 整体框线
    int style = 0;              // 0=单线, 1=粗线, 2=双线, 3=点线, 4=虚线
    int width = 15;             // 边框宽度（twips）
    int colorIndex = 0;         // 边框颜色索引
};

// ============================================================
// 段落格式状态
// ============================================================
struct ParaState {
    int align = 0;              // 对齐: 0=左, 1=居中, 2=右, 3=两端
    int leftIndent = 0;         // 左缩进（twips）
    int rightIndent = 0;        // 右缩进（twips）
    int firstIndent = 0;        // 首行缩进（twips）
    int spaceBefore = 0;        // 段前间距（twips）
    int spaceAfter = 0;         // 段后间距（twips）
    int lineSpacing = 0;        // 行间距（twips）
    int lineSpacingRule = 0;    // 行间距规则: 0=自动, 1=精确, 2=最小
    bool inTable = false;       // 是否在表格内
    int listSelector = 0;       // \ls 值
    int listLevel = 0;          // \ilvl 值
    int styleIndex = 0;         // 样式索引
    // 文字方向（Feature 9）
    bool rtlPar = false;        // \rtlpar 段落从右到左
    // 分页（Feature 11）
    bool pageBreakBefore = false; // \pagebb 段前分页
    // 段落边框（Feature 8）
    ParaBorder border;
    // 段落底纹（Feature 15）
    int shadingPct = 0;         // \shading 百分比
    int shadingFgColor = 0;     // \cfpat 前景色索引
    int shadingBgColor = 0;     // \cbpat 背景色索引
    // 列节（Feature 17）
    int outlineLevel = -1;      // \outlinelevel 值（0=H1, 1=H2...）
};

// ============================================================
// 目标类型
// ============================================================
enum class Destination {
    Normal,             // 普通内容
    FontTable,          // 字体表
    FontEntry,          // 字体条目
    ColorTable,         // 颜色表
    StyleSheet,         // 样式表
    StyleEntry,         // 样式条目
    Info,               // 文档信息
    Pict,               // 图片
    FieldInst,          // 域指令
    FieldResult,        // 域结果
    ListText,           // 列表文本
    ListTable,          // 列表表格
    ListEntry,          // 列表条目
    ListLevelEntry,     // 列表级别条目
    ListOverrideTable,  // 列表覆盖表
    ListOverrideEntry,  // 列表覆盖条目
    HtmlTag,            // \*\htmltag - Outlook 嵌入 HTML 标签
    HtmlRtf,            // \*\htmlrtf - 有 fromhtml 时跳过
    MhtmlTag,           // \*\mhtmltag
    Bookmark,           // 书签
    BookmarkStart,      // \bkmkstart - 书签开始（Feature 20）
    BookmarkEnd,        // \bkmkend - 书签结束（Feature 20）
    Shape,              // 图形
    ShapeInst,          // 图形实例
    ShapePict,          // 图形图片
    ShapeText,          // \shptxt - 图形文本框内容（Feature 13）
    Header,             // 页眉
    Footer,             // 页脚
    Footnote,           // 脚注（Feature 10）
    Endnote,            // 尾注（Feature 10）
    Skip,               // 未知可选目标 - 忽略
};

// ============================================================
// 样式表条目（Feature 2）
// ============================================================
struct StyleEntry {
    int index = 0;              // 样式索引
    std::string name;           // 样式名称
    bool isCharStyle = false;   // cs = 字符样式
    bool isParaStyle = false;   // s = 段落样式
    int basedOn = 0;            // \sbasedon 基础样式
    int next = 0;               // \snext 下一样式
    int outlineLevel = -1;      // \outlinelevel（0=H1, 1=H2...）
};

// ============================================================
// RTF 解析状态（每个组一个，保存到栈中）
// ============================================================
struct RtfState {
    CharState cs;                       // 字符格式状态
    ParaState ps;                       // 段落格式状态
    Destination dest = Destination::Normal;  // 当前目标
    bool optionalDest = false;          // 是否在 \* 之后
    bool skipGroup = false;             // 整个组需要跳过
    int groupDepth = 0;                 // 组嵌套深度
};

} // namespace rtf2html
