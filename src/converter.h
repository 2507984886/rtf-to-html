#pragma once
// converter.h - RTF 到 HTML 转换器接口
// 核心转换引擎，处理 RTF 状态机和 HTML 生成

#include "state.h"
#include "tokenizer.h"
#include "../include/rtf2html/rtf2html.h"
#include <string>
#include <vector>
#include <map>
#include <stack>
#include <sstream>
#include <cstdint>

namespace rtf2html {

// ============================================================
// 列表覆盖条目（\listoverride）
// ============================================================
struct ListOverride {
    int listId = 0;   // 对应的列表 ID
    int ls = 0;       // \ls 值
};

// ============================================================
// 颜色解析临时状态
// ============================================================
struct ColorParseTemp_ {
    int r = 0, g = 0, b = 0;
    bool hasR = false, hasG = false, hasB = false;
};

// ============================================================
// RTF 转换器
// ============================================================
class Converter {
public:
    /**
     * 构造函数
     * @param data  RTF 数据指针
     * @param size  数据长度
     * @param opts  转换选项
     */
    Converter(const char* data, size_t size, const ConvertOptions& opts);

    /**
     * 执行转换
     * @return 转换结果
     */
    ConvertResult convert();

private:
    const char* data_;           // RTF 数据
    size_t size_;                // 数据长度
    const ConvertOptions& opts_; // 转换选项

    // -------- 状态机 --------
    std::stack<RtfState> stateStack_;   // RTF 状态栈
    RtfState curState_;                  // 当前状态

    // -------- 文档级设置 --------
    int docCodepage_ = 1252;    // 文档 ANSI 代码页
    bool fromHtml_ = false;     // 是否是 Outlook fromhtml 模式

    // -------- 字体/颜色表 --------
    std::map<int, FontEntry> fontTable_;    // 字体表
    std::vector<ColorEntry> colorTable_;    // 颜色表
    ColorParseTemp_ colorTemp_;             // 颜色解析临时状态
    int currentFontIndex_ = -1;             // 正在解析的字体索引
    FontEntry currentFontEntry_;            // 正在解析的字体条目

    // -------- 列表表格 --------
    std::map<int, ListDef> listTable_;              // 列表定义表
    std::map<int, ListOverride> listOverrideTable_; // 列表覆盖表
    int currentListId_ = 0;                         // 正在解析的列表 ID
    ListDef currentListDef_;                        // 正在解析的列表定义
    int currentListLevel_ = 0;                      // 正在解析的列表级别
    ListOverride currentOverride_;                  // 正在解析的列表覆盖

    // -------- 样式表（Feature 2）--------
    std::map<int, StyleEntry> styleTable_;          // 样式表
    StyleEntry currentStyle_;                       // 正在解析的样式条目
    bool parsingStyle_ = false;                     // 是否在解析样式条目
    std::string styleNameBuffer_;                   // 样式名称缓冲

    // -------- DBCS 状态（Feature 1）--------
    uint8_t dbcsLeadByte_ = 0;          // DBCS 前导字节缓冲（0表示无待处理）

    // -------- 非组形式 htmlrtf（Feature 3）--------
    bool htmlRtfInlineSkip_ = false;    // 非组形式 \htmlrtf 跳过标志

    // -------- 脚注/尾注（Feature 10）--------
    std::vector<std::string> footnotes_;    // 收集的脚注 HTML
    std::vector<std::string> endnotes_;     // 收集的尾注 HTML
    std::string footnoteBuffer_;            // 正在收集的脚注内容
    int footnoteCounter_ = 0;              // 脚注计数器
    bool collectingFootnote_ = false;       // 是否正在收集脚注
    bool inFootnote_ = false;              // 是否在脚注组内

    // -------- 书签（Feature 20）--------
    std::string bookmarkNameBuffer_;        // 书签名称缓冲

    // -------- 段节（Feature 17）--------
    int sectionCols_ = 0;                  // 当前节的列数
    bool inSection_ = false;               // 是否在多列节中

    // -------- 图片状态 --------
    struct PictState {
        std::string hexData;          // 十六进制数据
        std::vector<uint8_t> binData; // 二进制数据
        std::string format;           // "png", "jpeg", "wmf", "emf", "bmp"
        int width = 0;                // 原始宽度（twips/像素）
        int height = 0;               // 原始高度（twips/像素）
        int widthGoal = 0;            // 目标宽度（twips）
        int heightGoal = 0;           // 目标高度（twips）
        int scaleX = 100;             // X 缩放比（百分比）
        int scaleY = 100;             // Y 缩放比（百分比）
        bool collecting = false;      // 是否正在收集数据
    } pict_;

    // -------- 域（Field）状态 --------
    struct FieldState {
        std::string instText;    // \fldinst 文本
        std::string resultHtml;  // \fldrslt HTML（未用）
        bool inInst = false;     // 是否在 \fldinst
        bool inResult = false;   // 是否在 \fldrslt
        std::string hyperlink;   // 提取的超链接 URL
    } field_;

    bool fieldAnchorOpen_ = false;  // 是否有未关闭的 <a> 标签

    // -------- 列表输出状态 --------
    struct ListOutputState {
        int ls = 0;
        int level = 0;
        bool isOrdered = false;
    };
    std::vector<ListOutputState> listStack_; // 列表嵌套栈

    // -------- 表格状态 --------
    struct TableState {
        bool active = false;                    // 是否在表格中
        bool rowDef = false;                    // 是否在定义行
        std::vector<TableCellDef> cellDefs;     // 当前行单元格定义
        TableCellDef currentCellDef;            // 正在定义的单元格
        std::vector<std::string> cellContents;  // 当前行各单元格内容
        std::string currentCellContent;         // 当前单元格内容
        bool inCell = false;                    // 是否在单元格内
        int cellIndex = 0;                      // 当前单元格索引
        // 表格对齐（Feature 16）
        int align = 0;                          // 0=左, 1=居中, 2=右
        // 嵌套支持（Feature 5）
        int itap = 1;                           // \itap 值：表格嵌套层级
    } table_;

    // 嵌套表格栈（Feature 5）
    std::vector<TableState> tableStack_;        // 嵌套表格状态栈

    // -------- HTML 输出 --------
    std::ostringstream html_;       // 主 HTML 输出流
    bool paraOpen_ = false;         // 是否有未关闭的段落
    std::string currentParaTag_ = "p"; // 当前段落标签名（p/h1/h2 等）（Feature 2）
    std::string currentSpanCss_;    // 当前打开的 span 的 CSS
    bool spanOpen_ = false;         // 是否有打开的 span

    // -------- HtmlTag 收集（fromhtml 模式） --------
    std::string htmlTagBuffer_;         // 正在收集的 htmltag 内容
    bool collectingHtmlTag_ = false;    // 是否正在收集

    // -------- 文本缓冲 --------
    std::string listTextBuffer_;    // \listtext 内容
    bool inListText_ = false;

    // -------- Unicode 跳过计数 --------
    int unicodeSkipCount_ = 0;  // 还需要跳过的 ANSI 字符数

    // ---- 内部方法 ----
    void processToken(const Token& tok);

    void handleGroupStart();
    void handleGroupEnd();
    void handleControlWord(const Token& tok);
    void handleControlSymbol(const Token& tok);
    void handleHexChar(const Token& tok);
    void handleText(const Token& tok);

    void handleFontTableControl(const std::string& name, bool hasParam, int param);
    void handleColorTableControl(const std::string& name, bool hasParam, int param);
    void handleCharFormatControl(const std::string& name, bool hasParam, int param);
    void handleParaFormatControl(const std::string& name, bool hasParam, int param);
    void handleTableControl(const std::string& name, bool hasParam, int param);
    void handleListControl(const std::string& name, bool hasParam, int param);
    void handlePictControl(const std::string& name, bool hasParam, int param);
    void handleSpecialOutput(const std::string& name);

    void openParagraph();
    void closeParagraph();
    void emitPar();
    void emitCell();
    void emitRow();
    void startTable();
    void endTable();
    void flushRow();

    void ensureSpanClosed();
    void ensureSpanOpen();
    std::string buildSpanCss() const;
    std::string buildParaCss() const;

    void openList(int ls, int level, bool ordered);
    void closeListsToLevel(int targetLevel);
    void closeAllLists();

    void emitPict();
    static std::string toBase64(const std::vector<uint8_t>& data);
    static std::vector<uint8_t> dibToPng(const std::vector<uint8_t>& dibData); // Feature 7
    static std::vector<uint8_t> hexToBinary(const std::string& hex);

    static std::string extractHyperlink(const std::string& instText);
    static std::string extractFieldContent(const std::string& instText,
                                            std::string& fieldType); // Feature 12

    // 脚注方法（Feature 10）
    void emitFootnotes();

    void writeToCurrentOutput(const std::string& s);

    static std::string htmlEscape(const std::string& text);
    static std::string htmlEscapeAttr(const std::string& text);

    // twips 转像素（96 dpi, 1440 twips/inch）
    static double twipsToPx(int twips) { return twips / 15.0; }

    std::string colorToHex(int index) const;
    std::string highlightToColor(int index) const;
    std::string getFontName(int index) const;
    void updateFontCodepage(int fontIndex);
};

} // namespace rtf2html
