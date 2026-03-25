// converter.cpp - RTF 到 HTML 转换器核心实现
// 处理 RTF 状态机、字体/颜色表、表格、图片、列表、fromhtml 模式
// 支持 DBCS 多字节编码、样式表、脚注、书签等高级特性

#include "converter.h"
#include "codepage.h"
#include <algorithm>
#include <cctype>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <iomanip>
#include <sstream>
#include <stdexcept>

namespace rtf2html {

// ============================================================
// Base64 编码表
// ============================================================
static const char kBase64Chars[] =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    "abcdefghijklmnopqrstuvwxyz"
    "0123456789+/";

// ============================================================
// 颜色表解析辅助状态（模块级，每次转换复用）
// ============================================================
struct ColorParseTemp {
    int r = 0, g = 0, b = 0;
    bool hasR = false, hasG = false, hasB = false;
};

// ============================================================
// 构造函数
// ============================================================
Converter::Converter(const char* data, size_t size, const ConvertOptions& opts)
    : data_(data), size_(size), opts_(opts)
{
    // 初始化默认颜色表（第 0 项为自动颜色）
    ColorEntry autoColor;
    autoColor.isAuto = true;
    colorTable_.push_back(autoColor);
}

// ============================================================
// 主转换入口
// ============================================================
ConvertResult Converter::convert() {
    ConvertResult result;

    try {
        Tokenizer tokenizer(data_, size_);

        // 输出 HTML 头部（如果需要完整页面）
        if (opts_.generateFullPage) {
            html_ << "<!DOCTYPE html>\n<html>\n<head>\n"
                  << "<meta charset=\"UTF-8\">\n"
                  << "<style>body{font-family:"
                  << opts_.defaultFontFamily
                  << ";font-size:" << opts_.defaultFontSizePt
                  << "pt;margin:8px;}</style>\n"
                  << "</head>\n<body>\n";
        }

        // 主解析循环
        while (!tokenizer.eof()) {
            Token tok = tokenizer.next();
            if (tok.type == TokenType::END_OF_INPUT) break;
            processToken(tok);
        }

        // 关闭未关闭的元素
        ensureSpanClosed();
        if (!fromHtml_) {
            closeParagraph();
            if (table_.active) endTable();
            closeAllLists();
        }

        // 关闭多列节（Feature 17）
        if (inSection_) {
            html_ << "</div>\n";
            inSection_ = false;
        }

        // 输出脚注区（Feature 10）
        emitFootnotes();

        if (opts_.generateFullPage) {
            html_ << "\n</body>\n</html>";
        }


        result.html = html_.str();
        result.success = true;
    } catch (const std::exception& e) {
        result.success = false;
        result.errorMessage = e.what();
        result.html = html_.str();
    }

    return result;
}

// ============================================================
// Token 分发处理
// ============================================================
void Converter::processToken(const Token& tok) {
    switch (tok.type) {
        case TokenType::GROUP_START:
            handleGroupStart();
            break;
        case TokenType::GROUP_END:
            handleGroupEnd();
            break;
        case TokenType::CONTROL_WORD:
            handleControlWord(tok);
            break;
        case TokenType::CONTROL_SYMBOL:
            handleControlSymbol(tok);
            break;
        case TokenType::HEX_CHAR:
            handleHexChar(tok);
            break;
        case TokenType::TEXT:
            handleText(tok);
            break;
        case TokenType::BINARY_DATA:
            if (curState_.dest == Destination::Pict) {
                pict_.binData.insert(pict_.binData.end(),
                    tok.binaryData.begin(), tok.binaryData.end());
            }
            break;
        case TokenType::END_OF_INPUT:
            break;
    }
}

// ============================================================
// 处理组开始 {
// ============================================================
void Converter::handleGroupStart() {
    stateStack_.push(curState_);
    curState_.groupDepth++;
    curState_.optionalDest = false;
    // 重置 skipGroup（子组不继承跳过状态，除非父组是 Skip 目标）
    if (curState_.dest == Destination::Skip ||
        curState_.dest == Destination::Info ||
        curState_.dest == Destination::HtmlRtf ||
        curState_.dest == Destination::MhtmlTag ||
        curState_.dest == Destination::Header ||
        curState_.dest == Destination::Footer) {
        curState_.skipGroup = true;
    }
    // 在脚注组内，子组也保持脚注状态
    if (curState_.dest == Destination::Footnote ||
        curState_.dest == Destination::Endnote) {
        // 不重置，保持继续收集脚注内容
    }
}

// ============================================================
// 处理组结束 }
// ============================================================
void Converter::handleGroupEnd() {
    Destination leavingDest = curState_.dest;

    // 离开特定目标时的收尾工作
    switch (leavingDest) {
        case Destination::FontEntry:
            if (currentFontIndex_ >= 0) {
                fontTable_[currentFontIndex_] = currentFontEntry_;
            }
            break;

        case Destination::Pict:
            if (!curState_.skipGroup) {
                emitPict();
            }
            pict_ = PictState{};
            break;

        case Destination::FieldResult: {
            // 域结果结束
            field_.inResult = false;
            break;
        }

        case Destination::FieldInst: {
            // 分析域内容，支持多种域类型（Feature 12）
            std::string fieldType;
            std::string extracted = extractFieldContent(field_.instText, fieldType);
            field_.inInst = false;
            if (fieldType == "HYPERLINK") {
                field_.hyperlink = extracted;
                if (!field_.hyperlink.empty() && !fromHtml_) {
                    if (!paraOpen_) openParagraph();
                    ensureSpanClosed();
                    // 检查是否有 \l 锚点（本地链接）
                    html_ << "<a href=\"" + htmlEscapeAttr(field_.hyperlink) + "\">";
                    fieldAnchorOpen_ = true;
                }
            } else if (fieldType == "REF") {
                // \REF 书签引用
                if (!fromHtml_ && !extracted.empty()) {
                    if (!paraOpen_) openParagraph();
                    html_ << "<a href=\"#bkmk_" << htmlEscapeAttr(extracted) << "\">";
                    fieldAnchorOpen_ = true;
                }
            } else if (fieldType == "INCLUDEPICTURE") {
                // 嵌入图片
                if (!fromHtml_ && !extracted.empty()) {
                    std::string imgTag = "<img src=\"" + htmlEscapeAttr(extracted) + "\" alt=\"\">";
                    writeToCurrentOutput(imgTag);
                }
            } else if (fieldType == "MAILTO") {
                if (!fromHtml_ && !extracted.empty()) {
                    if (!paraOpen_) openParagraph();
                    html_ << "<a href=\"mailto:" << htmlEscapeAttr(extracted) << "\">";
                    fieldAnchorOpen_ = true;
                }
            }
            // 其他域类型（DATE, TIME, PAGE 等）在 fldrslt 中处理
            break;
        }

        case Destination::HtmlTag:
            if (fromHtml_ && collectingHtmlTag_) {
                html_ << htmlTagBuffer_;
                htmlTagBuffer_.clear();
                collectingHtmlTag_ = false;
            }
            break;

        case Destination::ListEntry:
            if (currentListDef_.listId != 0) {
                listTable_[currentListDef_.listId] = currentListDef_;
            }
            break;

        case Destination::StyleEntry:
            // 样式条目解析完毕，存储到样式表（Feature 2）
            if (!currentStyle_.name.empty() || currentStyle_.index > 0) {
                styleTable_[currentStyle_.index] = currentStyle_;
            }
            parsingStyle_ = false;
            styleNameBuffer_.clear();
            break;

        case Destination::Footnote: {
            // 脚注内容收集完毕（Feature 10）
            int n = footnoteCounter_;
            std::string fnEntry = "<div id=\"fn_" + std::to_string(n) + "\">"
                + "<sup>" + std::to_string(n) + "</sup> "
                + footnoteBuffer_
                + "</div>\n";
            footnotes_.push_back(fnEntry);
            footnoteBuffer_.clear();
            collectingFootnote_ = false;
            inFootnote_ = false;
            break;
        }

        case Destination::Endnote: {
            // 尾注内容收集完毕（Feature 10）
            int n = (int)endnotes_.size() + 1;
            std::string fnEntry = "<div id=\"en_" + std::to_string(n) + "\">"
                + "<sup>" + std::to_string(n) + "</sup> "
                + footnoteBuffer_
                + "</div>\n";
            endnotes_.push_back(fnEntry);
            footnoteBuffer_.clear();
            collectingFootnote_ = false;
            break;
        }

        case Destination::BookmarkStart: {
            // 书签开始（Feature 20）
            if (!bookmarkNameBuffer_.empty() && !fromHtml_) {
                // 输出锚点标签
                std::string anchor = "<a id=\"bkmk_" + htmlEscapeAttr(bookmarkNameBuffer_) + "\"></a>";
                if (paraOpen_ || (curState_.ps.inTable && table_.active)) {
                    writeToCurrentOutput(anchor);
                } else {
                    html_ << anchor;
                }
                bookmarkNameBuffer_.clear();
            }
            break;
        }

        default:
            break;
    }

    // 如果离开 field 组（field 是外层组）
    if (leavingDest == Destination::Normal && fieldAnchorOpen_) {
        // 检查是否需要关闭 <a>
        // 通过检查栈深度来判断
        // 简化：依赖 \fldrslt 来关闭
    }

    // 恢复状态
    if (!stateStack_.empty()) {
        RtfState prev = stateStack_.top();
        stateStack_.pop();

        // 如果离开了 fldrslt 且有超链接
        if (leavingDest == Destination::FieldResult && fieldAnchorOpen_) {
            ensureSpanClosed();
            html_ << "</a>";
            fieldAnchorOpen_ = false;
            field_ = FieldState{};
        }

        curState_ = prev;
    }
}

// ============================================================
// 处理控制字
// ============================================================
void Converter::handleControlWord(const Token& tok) {
    const std::string& name = tok.name;
    bool hasParam = tok.hasParam;
    int param = tok.param;

    // 跳过状态
    if (curState_.skipGroup) return;

    // ---- 文档级控制字 ----
    if (name == "rtf") return;

    if (name == "ansicpg") {
        docCodepage_ = ansicpgToCodepage(param);
        curState_.cs.fontCodepage = docCodepage_;
        return;
    }
    if (name == "fromhtml") {
        if (param == 1) fromHtml_ = true;
        return;
    }
    if (name == "deff") {
        curState_.cs.fontIndex = param;
        updateFontCodepage(param);
        return;
    }
    if (name == "deflang" || name == "deflangfe") return;

    // ---- 目标切换控制字 ----
    if (name == "fonttbl") {
        curState_.dest = Destination::FontTable;
        return;
    }
    if (name == "colortbl") {
        colorTable_.clear();
        ColorEntry autoColor; autoColor.isAuto = true;
        colorTable_.push_back(autoColor);
        curState_.dest = Destination::ColorTable;
        colorTemp_ = ColorParseTemp_{};
        return;
    }
    if (name == "stylesheet") {
        // Feature 2：样式表解析（不跳过，改为解析）
        curState_.dest = Destination::StyleSheet;
        return;
    }
    if (name == "info") {
        curState_.dest = Destination::Info;
        curState_.skipGroup = true;
        return;
    }
    if (name == "pict") {
        curState_.dest = Destination::Pict;
        pict_ = PictState{};
        pict_.collecting = true;
        return;
    }
    if (name == "fldinst") {
        curState_.dest = Destination::FieldInst;
        field_.instText.clear();
        field_.inInst = true;
        return;
    }
    if (name == "fldrslt") {
        curState_.dest = Destination::FieldResult;
        field_.inResult = true;
        return;
    }
    if (name == "field") {
        field_ = FieldState{};
        fieldAnchorOpen_ = false;
        return;
    }
    if (name == "listtext" || name == "pntext") {
        curState_.dest = Destination::ListText;
        listTextBuffer_.clear();
        inListText_ = true;
        return;
    }
    if (name == "listtable") {
        curState_.dest = Destination::ListTable;
        return;
    }
    if (name == "list") {
        if (curState_.dest == Destination::ListTable) {
            curState_.dest = Destination::ListEntry;
            currentListDef_ = ListDef{};
        }
        return;
    }
    if (name == "listlevel") {
        if (curState_.dest == Destination::ListEntry) {
            curState_.dest = Destination::ListLevelEntry;
            currentListDef_.levels.emplace_back();
            currentListLevel_ = (int)currentListDef_.levels.size() - 1;
        }
        return;
    }
    if (name == "listoverridetable") {
        curState_.dest = Destination::ListOverrideTable;
        return;
    }
    if (name == "listoverride") {
        if (curState_.dest == Destination::ListOverrideTable) {
            curState_.dest = Destination::ListOverrideEntry;
            currentOverride_ = ListOverride{};
        }
        return;
    }
    if (name == "htmltag") {
        if (fromHtml_) {
            curState_.dest = Destination::HtmlTag;
            htmlTagBuffer_.clear();
            collectingHtmlTag_ = true;
        } else {
            curState_.dest = Destination::Skip;
            curState_.skipGroup = true;
        }
        return;
    }
    if (name == "htmlrtf") {
        // Feature 3：支持组形式和非组形式 htmlrtf
        if (fromHtml_) {
            if (curState_.optionalDest) {
                // 组形式：{\*\htmlrtf ...}
                curState_.dest = Destination::HtmlRtf;
                curState_.skipGroup = true;
            } else if (hasParam && param == 0) {
                // \htmlrtf0 - 结束非组跳过
                htmlRtfInlineSkip_ = false;
            } else {
                // \htmlrtf（无参数或参数!=0）- 开始非组跳过
                htmlRtfInlineSkip_ = true;
            }
        }
        return;
    }
    if (name == "mhtmltag") {
        curState_.dest = Destination::MhtmlTag;
        curState_.skipGroup = true;
        return;
    }
    if (name == "shp" || name == "shpinst" || name == "shppict" ||
        name == "doshpict") {
        // 图形目标：只处理 shppict 内部的 pict
        curState_.dest = Destination::Shape;
        return;
    }
    if (name == "header"  || name == "footer" ||
        name == "headerl" || name == "headerr" || name == "headerf" ||
        name == "footerl" || name == "footerr" || name == "footerf" ||
        name == "headerrl"|| name == "footerrl") {
        curState_.dest = Destination::Skip;
        curState_.skipGroup = true;
        return;
    }
    if (name == "bkmkstart") {
        // Feature 20：书签开始
        curState_.dest = Destination::BookmarkStart;
        bookmarkNameBuffer_.clear();
        return;
    }
    if (name == "bkmkend") {
        // Feature 20：书签结束（不需要输出）
        curState_.dest = Destination::BookmarkEnd;
        return;
    }
    if (name == "footnote") {
        // Feature 10：脚注
        ++footnoteCounter_;
        // 在当前位置输出引用标记
        std::string ref = "<sup><a id=\"fnref_" + std::to_string(footnoteCounter_) + "\""
            + " href=\"#fn_" + std::to_string(footnoteCounter_) + "\">"
            + "[" + std::to_string(footnoteCounter_) + "]</a></sup>";
        if (!fromHtml_) {
            if (!paraOpen_) openParagraph();
            ensureSpanClosed();
            html_ << ref;
        }
        curState_.dest = Destination::Footnote;
        footnoteBuffer_.clear();
        collectingFootnote_ = true;
        inFootnote_ = true;
        return;
    }
    if (name == "xe" || name == "tc") {
        // 索引条目/目录条目，忽略
        curState_.dest = Destination::Skip;
        curState_.skipGroup = true;
        return;
    }

    // 未知可选目标（\* 之后）
    if (curState_.optionalDest) {
        static const std::string knownDests[] = {
            "fonttbl","colortbl","stylesheet","info","pict",
            "fldinst","fldrslt","listtext","pntext","listtable",
            "list","listlevel","listoverridetable","listoverride",
            "htmltag","htmlrtf","mhtmltag","shp","shpinst","shppict",
            "header","footer","bkmkstart","bkmkend","field",
            "headerl","headerr","footerl","footerr","doshpict",
            "footnote","xe","tc","shptxt",
            ""
        };
        bool known = false;
        for (const auto& kd : knownDests) {
            if (kd.empty()) break;
            if (name == kd) { known = true; break; }
        }
        if (!known) {
            curState_.dest = Destination::Skip;
            curState_.skipGroup = true;
            curState_.optionalDest = false;
            return;
        }
    }

    // 根据当前目标分发处理
    switch (curState_.dest) {
        case Destination::FontTable:
        case Destination::FontEntry:
            handleFontTableControl(name, hasParam, param);
            return;
        case Destination::ColorTable:
            handleColorTableControl(name, hasParam, param);
            return;
        case Destination::ListTable:
        case Destination::ListEntry:
        case Destination::ListLevelEntry:
        case Destination::ListOverrideTable:
        case Destination::ListOverrideEntry:
            handleListControl(name, hasParam, param);
            return;
        case Destination::Pict:
            handlePictControl(name, hasParam, param);
            return;
        case Destination::FieldInst:
            // 字段指令中忽略格式控制
            return;
        case Destination::Skip:
        case Destination::Info:
        case Destination::HtmlRtf:
        case Destination::MhtmlTag:
        case Destination::BookmarkEnd:
            // 跳过这些目标中的控制字
            return;
        case Destination::StyleSheet:
            // Feature 2：样式表内，当遇到新 { 时创建子条目
            if (name == "s" && hasParam) {
                // 段落样式条目开始
                curState_.dest = Destination::StyleEntry;
                currentStyle_ = StyleEntry{};
                currentStyle_.index = param;
                currentStyle_.isParaStyle = true;
                parsingStyle_ = true;
                styleNameBuffer_.clear();
            } else if (name == "cs" && hasParam) {
                // 字符样式条目开始
                curState_.dest = Destination::StyleEntry;
                currentStyle_ = StyleEntry{};
                currentStyle_.index = param;
                currentStyle_.isCharStyle = true;
                parsingStyle_ = true;
                styleNameBuffer_.clear();
            }
            return;
        case Destination::StyleEntry:
            // Feature 2：解析样式条目中的控制字
            if (name == "sbasedon" && hasParam) {
                currentStyle_.basedOn = param;
            } else if (name == "snext" && hasParam) {
                currentStyle_.next = param;
            } else if (name == "outlinelevel" && hasParam) {
                currentStyle_.outlineLevel = param;
            } else if (name == "cs" && hasParam) {
                // 字符样式索引（在段落样式内）
                currentStyle_.isCharStyle = true;
            }
            // 样式条目内的格式控制字可以忽略（仅存储 outlinelevel）
            return;
        case Destination::ListText:
            // 跳过列表文本中的控制字
            return;
        default:
            break;
    }

    // ---- 正常目标及 FieldResult 中的控制字 ----
    handleCharFormatControl(name, hasParam, param);
    handleParaFormatControl(name, hasParam, param);
    handleTableControl(name, hasParam, param);
    handleSpecialOutput(name);
}

// ============================================================
// 字体表控制字处理
// ============================================================
void Converter::handleFontTableControl(const std::string& name, bool /*hasParam*/, int param) {
    if (name == "f") {
        // 新字体条目开始
        curState_.dest = Destination::FontEntry;
        currentFontIndex_ = param;
        currentFontEntry_ = FontEntry{};
        return;
    }
    if (curState_.dest != Destination::FontEntry) return;

    if (name == "fcharset") {
        currentFontEntry_.charset = param;
        currentFontEntry_.codepage = fcharsetToCodepage(param);
        return;
    }
    if (name == "fcodepage") {
        currentFontEntry_.codepage = param;
        return;
    }
    if (name == "froman")  { currentFontEntry_.family = 1; return; }
    if (name == "fswiss")  { currentFontEntry_.family = 2; return; }
    if (name == "fmodern") { currentFontEntry_.family = 3; return; }
    if (name == "fscript") { currentFontEntry_.family = 4; return; }
    if (name == "fdecor")  { currentFontEntry_.family = 5; return; }
    if (name == "ftech")   { currentFontEntry_.family = 6; return; }
    if (name == "fbidi")   { currentFontEntry_.family = 7; return; }
    if (name == "fnil")    { currentFontEntry_.family = 0; return; }
    if (name == "cpg") {
        currentFontEntry_.codepage = param;
        return;
    }
}

// ============================================================
// 颜色表控制字处理
// ============================================================
void Converter::handleColorTableControl(const std::string& name, bool /*hasParam*/, int param) {
    if (name == "red")   { colorTemp_.r = param; colorTemp_.hasR = true; return; }
    if (name == "green") { colorTemp_.g = param; colorTemp_.hasG = true; return; }
    if (name == "blue")  { colorTemp_.b = param; colorTemp_.hasB = true; return; }
}

// ============================================================
// 字符格式控制字处理
// ============================================================
void Converter::handleCharFormatControl(const std::string& name, bool hasParam, int param) {
    CharState& cs = curState_.cs;

    // 在改变格式前关闭当前 span
    auto maybeCloseSpan = [&]() {
        if (spanOpen_ && !fromHtml_) {
            ensureSpanClosed();
        }
    };

    if (name == "f") {
        maybeCloseSpan();
        cs.fontIndex = param;
        updateFontCodepage(param);
        return;
    }
    if (name == "fs") {
        maybeCloseSpan();
        cs.fontSize = hasParam ? param : 24;
        return;
    }
    if (name == "b") {
        maybeCloseSpan();
        cs.bold = !hasParam || param != 0;
        return;
    }
    if (name == "i") {
        maybeCloseSpan();
        cs.italic = !hasParam || param != 0;
        return;
    }
    if (name == "ul") {
        maybeCloseSpan();
        cs.underline = !hasParam || param != 0;
        cs.underlineStyle = cs.underline ? 1 : 0;
        return;
    }
    if (name == "uld"  || name == "uldash"  || name == "uldb" ||
        name == "ulhwave" || name == "ulldash" || name == "ulth" ||
        name == "ulthd"|| name == "ulthdash"|| name == "ulwave"||
        name == "ulw") {
        maybeCloseSpan();
        cs.underline = true; cs.underlineStyle = 2; return;
    }
    if (name == "ulnone") {
        maybeCloseSpan();
        cs.underline = false; cs.underlineStyle = 0; return;
    }
    if (name == "strike" || name == "striked") {
        maybeCloseSpan();
        cs.strikethrough = !hasParam || param != 0; return;
    }
    if (name == "striked1") {
        maybeCloseSpan();
        cs.strikethrough = true; return;
    }
    if (name == "super") {
        maybeCloseSpan();
        cs.superscript = !hasParam || param != 0;
        cs.subscript = false; return;
    }
    if (name == "sub") {
        maybeCloseSpan();
        cs.subscript = !hasParam || param != 0;
        cs.superscript = false; return;
    }
    if (name == "nosupersub") {
        maybeCloseSpan();
        cs.superscript = false; cs.subscript = false; return;
    }
    if (name == "v") {
        maybeCloseSpan();
        cs.hidden = !hasParam || param != 0; return;
    }
    if (name == "caps") {
        maybeCloseSpan();
        cs.allCaps = !hasParam || param != 0; return;
    }
    if (name == "scaps") {
        maybeCloseSpan();
        cs.smallCaps = !hasParam || param != 0; return;
    }
    if (name == "cf") {
        maybeCloseSpan();
        cs.fgColorIndex = param; return;
    }
    if (name == "cb" || name == "chcbpat") {
        maybeCloseSpan();
        cs.bgColorIndex = param; return;
    }
    if (name == "highlight") {
        maybeCloseSpan();
        cs.highlightIndex = param; return;
    }
    if (name == "expnd" || name == "expndtw") {
        maybeCloseSpan();
        cs.charSpacing = param; return;
    }
    if (name == "charscalex") {
        maybeCloseSpan();
        cs.charScaleX = hasParam ? param : 100; return;
    }
    if (name == "lang" || name == "langfe" || name == "langnp") {
        cs.lang = param; return;
    }
    if (name == "plain") {
        maybeCloseSpan();
        int uc = cs.unicodeSkip;
        int cp = cs.fontCodepage;
        cs = CharState{};
        cs.unicodeSkip = uc;
        cs.fontCodepage = cp;
        return;
    }
    if (name == "uc") {
        cs.unicodeSkip = hasParam ? param : 1; return;
    }
    if (name == "u") {
        // Unicode 字符
        int cp = param;
        if (cp < 0) cp += 65536;
        unicodeSkipCount_ = cs.unicodeSkip;
        std::string utf8 = utf32ToUtf8(static_cast<uint32_t>(cp));
        writeToCurrentOutput(htmlEscape(utf8));
        return;
    }
    if (name == "loch" || name == "hich" || name == "dbch") {
        // 字体切换（低字节/高字节/双字节），忽略
        return;
    }
    if (name == "cs") {
        cs.styleIndex = param; return;
    }
    // Feature 14：特殊文字效果
    if (name == "outl") {
        maybeCloseSpan();
        cs.outline = !hasParam || param != 0; return;
    }
    if (name == "outl0") {
        maybeCloseSpan();
        cs.outline = false; return;
    }
    if (name == "shad") {
        maybeCloseSpan();
        cs.shadow = !hasParam || param != 0; return;
    }
    if (name == "shad0") {
        maybeCloseSpan();
        cs.shadow = false; return;
    }
    if (name == "embo") {
        maybeCloseSpan();
        cs.emboss = !hasParam || param != 0; return;
    }
    if (name == "embo0") {
        maybeCloseSpan();
        cs.emboss = false; return;
    }
    if (name == "impr") {
        maybeCloseSpan();
        cs.engrave = !hasParam || param != 0; return;
    }
    if (name == "impr0") {
        maybeCloseSpan();
        cs.engrave = false; return;
    }
    // Feature 9：文字方向
    if (name == "rtlch") {
        maybeCloseSpan();
        cs.rtl = true; return;
    }
    if (name == "ltrch") {
        maybeCloseSpan();
        cs.rtl = false; return;
    }
}

// ============================================================
// 段落格式控制字处理
// ============================================================
void Converter::handleParaFormatControl(const std::string& name, bool hasParam, int param) {
    ParaState& ps = curState_.ps;

    if (name == "pard") {
        bool inTable = ps.inTable;
        int ls = ps.listSelector;
        int ilvl = ps.listLevel;
        ps = ParaState{};
        ps.inTable = inTable;
        // pard 清除列表选择器
        (void)ls; (void)ilvl;
        return;
    }
    if (name == "ql") { ps.align = 0; return; }
    if (name == "qc") { ps.align = 1; return; }
    if (name == "qr") { ps.align = 2; return; }
    if (name == "qj") { ps.align = 3; return; }
    if (name == "li") { ps.leftIndent = param; return; }
    if (name == "ri") { ps.rightIndent = param; return; }
    if (name == "fi") { ps.firstIndent = param; return; }
    if (name == "sb") { ps.spaceBefore = param; return; }
    if (name == "sa") { ps.spaceAfter = param; return; }
    if (name == "sl") { ps.lineSpacing = hasParam ? param : 0; return; }
    if (name == "slmult") { ps.lineSpacingRule = param; return; }
    if (name == "intbl") {
        ps.inTable = true;
        if (!table_.active) startTable();
        return;
    }
    if (name == "ls") {
        ps.listSelector = param;
        return;
    }
    if (name == "ilvl") {
        ps.listLevel = param;
        return;
    }
    if (name == "par") {
        emitPar();
        return;
    }
    if (name == "line") {
        ensureSpanClosed();
        if (curState_.ps.inTable && table_.active) {
            table_.currentCellContent += "<br>";
        } else {
            if (paraOpen_) html_ << "<br>";
            else html_ << "<br>";
        }
        return;
    }
    if (name == "page") {
        // Feature 11：分页符
        ensureSpanClosed();
        if (!fromHtml_) {
            closeParagraph();
            html_ << "<div style=\"page-break-after:always;border-top:1px dashed #ccc;margin:1em 0;\"></div>\n";
        }
        return;
    }
    if (name == "pagebb") {
        // Feature 11：段前分页
        ps.pageBreakBefore = true;
        return;
    }
    if (name == "sect") {
        // Feature 17：节分隔符
        if (!fromHtml_) {
            ensureSpanClosed();
            closeParagraph();
            if (inSection_) {
                html_ << "</div>\n";
                inSection_ = false;
            }
        }
        return;
    }
    if (name == "sectd") {
        // 节属性重置
        sectionCols_ = 0;
        return;
    }
    if (name == "cols") {
        // Feature 17：多列
        sectionCols_ = hasParam ? param : 1;
        if (!fromHtml_ && sectionCols_ > 1) {
            ensureSpanClosed();
            closeParagraph();
            if (inSection_) {
                html_ << "</div>\n";
            }
            html_ << "<div style=\"column-count:" << sectionCols_ << ";\">\n";
            inSection_ = true;
        }
        return;
    }
    if (name == "s") {
        ps.styleIndex = param;
        // Feature 2：应用样式表中的 outlinelevel
        if (!styleTable_.empty()) {
            auto it = styleTable_.find(param);
            if (it != styleTable_.end()) {
                ps.outlineLevel = it->second.outlineLevel;
            }
        }
        return;
    }
    if (name == "outlinelevel") {
        ps.outlineLevel = hasParam ? param : -1;
        return;
    }
    if (name == "nowidctlpar" || name == "widctlpar" ||
        name == "keepn" || name == "keep" ||
        name == "hyphpar" || name == "nohyph") {
        return;
    }
    // Feature 9：段落文字方向
    if (name == "rtlpar") {
        ps.rtlPar = true; return;
    }
    if (name == "ltrpar") {
        ps.rtlPar = false; return;
    }
    if (name == "rtlrow" || name == "ltrrow") {
        return; // 表格行方向（暂忽略）
    }
    // Feature 8：段落边框
    if (name == "brdrl") { ps.border.left = true; return; }
    if (name == "brdrr") { ps.border.right = true; return; }
    if (name == "brdrt") { ps.border.top = true; return; }
    if (name == "brdrb") { ps.border.bottom = true; return; }
    if (name == "box")   { ps.border.box = true; return; }
    if (name == "brdrs") { ps.border.style = 0; return; }   // 单线
    if (name == "brdrth"){ ps.border.style = 1; return; }   // 粗线
    if (name == "brdrdb"){ ps.border.style = 2; return; }   // 双线
    if (name == "brdrdot"){ ps.border.style = 3; return; }  // 点线
    if (name == "brdrdash"){ ps.border.style = 4; return; } // 虚线
    if (name == "brdrw") { ps.border.width = hasParam ? param : 15; return; }
    if (name == "brdrcf") { ps.border.colorIndex = hasParam ? param : 0; return; }
    // Feature 15：段落底纹
    if (name == "shading") { ps.shadingPct = hasParam ? param : 0; return; }
    if (name == "cfpat")   { ps.shadingFgColor = hasParam ? param : 0; return; }
    if (name == "cbpat")   { ps.shadingBgColor = hasParam ? param : 0; return; }
    // Feature 18：特殊字符
    if (name == "softline") {
        if (!fromHtml_) {
            if (curState_.ps.inTable && table_.active) {
                table_.currentCellContent += "<br>";
            } else {
                writeToCurrentOutput("<br>");
            }
        }
        return;
    }
    (void)hasParam;
}

// ============================================================
// 表格控制字处理
// ============================================================
void Converter::handleTableControl(const std::string& name, bool hasParam, int param) {
    if (name == "trowd") {
        if (!table_.active) startTable();
        table_.rowDef = true;
        table_.cellDefs.clear();
        table_.currentCellDef = TableCellDef{};
        return;
    }
    if (name == "cellx") {
        table_.currentCellDef.rightBoundary = param;
        table_.cellDefs.push_back(table_.currentCellDef);
        table_.currentCellDef = TableCellDef{};
        return;
    }
    if (name == "cell") {
        emitCell();
        return;
    }
    if (name == "row") {
        emitRow();
        return;
    }
    if (name == "nestrow") {
        emitRow();
        return;
    }
    if (name == "clcbpat" || name == "clcbpatraw") {
        table_.currentCellDef.bgColorIndex = param; return;
    }
    if (name == "clbrdrl") { table_.currentCellDef.borderLeft   = true; return; }
    if (name == "clbrdrr") { table_.currentCellDef.borderRight  = true; return; }
    if (name == "clbrdrt") { table_.currentCellDef.borderTop    = true; return; }
    if (name == "clbrdrb") { table_.currentCellDef.borderBottom = true; return; }
    if (name == "clvertalt") { table_.currentCellDef.valign = 0; return; }
    if (name == "clvertalc") { table_.currentCellDef.valign = 1; return; }
    if (name == "clvertalb") { table_.currentCellDef.valign = 2; return; }
    if (name == "clvmgf")  { table_.currentCellDef.vertMergeFirst = true; return; }
    if (name == "clvmrg")  { table_.currentCellDef.vertMerge = true; return; }
    if (name == "trrh")    { return; }
    if (name == "trhdr")   { return; }
    if (name == "trkeep")  { return; }
    if (name == "brdrw")   { table_.currentCellDef.borderLineWidth = param; return; }
    // Feature 16：表格对齐
    if (name == "trql")    { table_.align = 0; return; } // 左对齐
    if (name == "trqr")    { table_.align = 2; return; } // 右对齐
    if (name == "trqc")    { table_.align = 1; return; } // 居中
    // Feature 16：单元格内边距
    if (name == "clpadl" || name == "clpadr" ||
        name == "clpadt" || name == "clpadb") { return; } // 暂存储
    // Feature 5：嵌套表格
    if (name == "itap") {
        table_.itap = hasParam ? param : 1;
        return;
    }
    // 忽略其他表格控制字
    if (name == "trgaph" || name == "trwWidth" || name == "trftsWidth" ||
        name == "trbrdrl" || name == "trbrdrr" || name == "trbrdrt" ||
        name == "trbrdrb" || name == "trbrdrh" || name == "trbrdrv" ||
        name == "clshdng" || name == "clcfpat") {
        return;
    }
}

// ============================================================
// 列表控制字处理
// ============================================================
void Converter::handleListControl(const std::string& name, bool hasParam, int param) {
    (void)hasParam;
    if (name == "listid") {
        if (curState_.dest == Destination::ListEntry) {
            currentListDef_.listId = param;
        } else if (curState_.dest == Destination::ListOverrideEntry) {
            currentOverride_.listId = param;
        }
        return;
    }
    if (name == "ls") {
        if (curState_.dest == Destination::ListOverrideEntry) {
            currentOverride_.ls = param;
            // Feature 4：修复 listOverrideTable_ 存储逻辑
            listOverrideTable_[param] = currentOverride_;
        }
        return;
    }
    if (name == "levelnfc" || name == "levelnfcn") {
        // Feature 4：存储列表类型
        if (!currentListDef_.levels.empty()) {
            currentListDef_.levels.back().numType = param;
        }
        return;
    }
    if (name == "levelstartat") {
        if (!currentListDef_.levels.empty()) {
            currentListDef_.levels.back().start = param;
        }
        return;
    }
    if (name == "listtemplateid") return;
    // 列表级别进入时，已在 handleControlWord 主分支处理
}

// ============================================================
// 图片控制字处理
// ============================================================
void Converter::handlePictControl(const std::string& name, bool /*hasParam*/, int param) {
    if (name == "pngblip")   { pict_.format = "png"; return; }
    if (name == "jpegblip")  { pict_.format = "jpeg"; return; }
    if (name == "wmetafile") { pict_.format = "wmf"; return; }
    if (name == "emfblip")   { pict_.format = "emf"; return; }
    if (name == "picbmp" || name == "dibitmap" || name == "wbitmap") {
        pict_.format = "bmp"; return;
    }
    if (name == "picw")     { pict_.width = param; return; }
    if (name == "pich")     { pict_.height = param; return; }
    if (name == "picwgoal") { pict_.widthGoal = param; return; }
    if (name == "pichgoal") { pict_.heightGoal = param; return; }
    if (name == "picscalex") { pict_.scaleX = param; return; }
    if (name == "picscaley") { pict_.scaleY = param; return; }
    if (name == "bliptag" || name == "blipuid") { return; }
    // 其他图片属性忽略
}

// ============================================================
// 特殊字符输出控制字
// ============================================================
void Converter::handleSpecialOutput(const std::string& name) {
    if (name == "tab") {
        writeToCurrentOutput(
            "<span style=\"display:inline-block;width:2em\">\xc2\xa0</span>");
        return;
    }
    if (name == "emdash")     { writeToCurrentOutput("&mdash;"); return; }
    if (name == "endash")     { writeToCurrentOutput("&ndash;"); return; }
    if (name == "bullet")     { writeToCurrentOutput("&#x2022;"); return; }
    if (name == "lquote")     { writeToCurrentOutput("&lsquo;"); return; }
    if (name == "rquote")     { writeToCurrentOutput("&rsquo;"); return; }
    if (name == "ldblquote")  { writeToCurrentOutput("&ldquo;"); return; }
    if (name == "rdblquote")  { writeToCurrentOutput("&rdquo;"); return; }
    if (name == "enspace")    { writeToCurrentOutput("&ensp;"); return; }
    if (name == "emspace")    { writeToCurrentOutput("&emsp;"); return; }
    if (name == "qmspace")    { writeToCurrentOutput("&#x2005;"); return; }
    if (name == "zwj")        { writeToCurrentOutput("&#x200D;"); return; }
    if (name == "zwnj")       { writeToCurrentOutput("&#x200C;"); return; }
    if (name == "ltrmark")    { writeToCurrentOutput("&#x200E;"); return; }
    if (name == "rtlmark")    { writeToCurrentOutput("&#x200F;"); return; }
    if (name == "chpgn")      { writeToCurrentOutput("<span class=\"page-num\">[PAGE]</span>"); return; }
    // Feature 18：额外特殊字符
    if (name == "line")       { writeToCurrentOutput("<br>"); return; }
    if (name == "softcol")    { writeToCurrentOutput("<br>"); return; }
    if (name == "nobreak")    { writeToCurrentOutput("&nbsp;"); return; }
    if (name == "zwbo")       { writeToCurrentOutput("&#x200B;"); return; }
    if (name == "chdate")     { writeToCurrentOutput("<time>[DATE]</time>"); return; }
    if (name == "chtime")     { writeToCurrentOutput("<time>[TIME]</time>"); return; }
    if (name == "column")     { writeToCurrentOutput("<br>"); return; } // Feature 11
    // Feature 12：特殊域字段输出（当作为单独控制字出现时）
    if (name == "date")       { writeToCurrentOutput("<time>[DATE]</time>"); return; }
    if (name == "time")       { writeToCurrentOutput("<time>[TIME]</time>"); return; }
}

// ============================================================
// 处理控制符号
// ============================================================
void Converter::handleControlSymbol(const Token& tok) {
    if (curState_.skipGroup) return;

    char sym = tok.symbol;

    if (sym == '*') {
        curState_.optionalDest = true;
        return;
    }

    // 在颜色表中 ; 是分隔符（控制符号形式）
    if (sym == ';' && curState_.dest == Destination::ColorTable) {
        if (!colorTemp_.hasR && !colorTemp_.hasG && !colorTemp_.hasB) {
            // 自动颜色条目（已在初始化时添加过 index 0）
            if (colorTable_.size() > 1) {
                ColorEntry ce; ce.isAuto = true;
                colorTable_.push_back(ce);
            }
        } else {
            ColorEntry ce;
            ce.isAuto = false;
            ce.r = static_cast<uint8_t>(colorTemp_.r);
            ce.g = static_cast<uint8_t>(colorTemp_.g);
            ce.b = static_cast<uint8_t>(colorTemp_.b);
            colorTable_.push_back(ce);
        }
        colorTemp_ = ColorParseTemp_{};
        return;
    }

    switch (sym) {
        case '\\': writeToCurrentOutput("\\"); return;
        case '{':  writeToCurrentOutput("{"); return;
        case '}':  writeToCurrentOutput("}"); return;
        case '~':  writeToCurrentOutput("&nbsp;"); return;
        case '-':  return; // 可选连字符，忽略
        case '_':  writeToCurrentOutput("&#x2011;"); return; // 不换行连字符
        default:   return;
    }
}

// ============================================================
// 处理十六进制字符
// ============================================================
void Converter::handleHexChar(const Token& tok) {
    if (curState_.skipGroup) return;
    // Feature 3：非组形式 htmlrtf 跳过
    if (htmlRtfInlineSkip_) return;

    if (unicodeSkipCount_ > 0) {
        --unicodeSkipCount_;
        // DBCS 前导字节在跳过时也需要清除
        dbcsLeadByte_ = 0;
        return;
    }

    uint8_t byte = tok.hexByte;
    int cp = curState_.cs.fontCodepage > 0 ? curState_.cs.fontCodepage : docCodepage_;

    switch (curState_.dest) {
        case Destination::Pict:
            // 十六进制字符在 pict 中：已在 TEXT 中处理
            break;

        case Destination::FontEntry: {
            // 字体名称中的字符（字体名不涉及 DBCS）
            std::string ch = cpToUtf8(byte, cp);
            currentFontEntry_.name += ch;
            break;
        }

        case Destination::HtmlTag:
            if (fromHtml_ && collectingHtmlTag_) {
                htmlTagBuffer_ += cpToUtf8(byte, cp);
            }
            break;

        case Destination::FieldInst:
            field_.instText += cpToUtf8(byte, cp);
            break;

        case Destination::ListText:
            listTextBuffer_ += cpToUtf8(byte, cp);
            break;

        case Destination::Skip:
        case Destination::Info:
        case Destination::HtmlRtf:
        case Destination::MhtmlTag:
        case Destination::StyleSheet:
        case Destination::StyleEntry:
        case Destination::BookmarkEnd:
            break;

        case Destination::BookmarkStart:
            // Feature 20：收集书签名称
            if (!fromHtml_) {
                bookmarkNameBuffer_ += cpToUtf8(byte, cp);
            }
            break;

        case Destination::Footnote:
        case Destination::Endnote:
            // Feature 10：收集脚注内容
            if (collectingFootnote_) {
                footnoteBuffer_ += htmlEscape(cpToUtf8(byte, cp));
            }
            break;

        default: {
            // Feature 1：DBCS 多字节编码处理
            // 注意：必须先检查 dbcsLeadByte_ != 0，因为 GBK/Big5 的跟随字节
            // 范围与前导字节范围有重叠（如 GBK 跟随字节 0xE3 也在前导字节范围内）
            if (dbcsLeadByte_ != 0) {
                // 已有前导字节在缓存，当前字节是跟随字节
                std::string decoded = dbcsPairToUtf8(dbcsLeadByte_, byte, cp);
                dbcsLeadByte_ = 0;
                writeToCurrentOutput(htmlEscape(decoded));
            } else if (isDbcsLeadByte(byte, cp)) {
                // 这是前导字节，缓存它，等待跟随字节
                dbcsLeadByte_ = byte;
            } else {
                // 单字节字符
                std::string decoded = cpToUtf8(byte, cp);
                writeToCurrentOutput(htmlEscape(decoded));
            }
            break;
        }
    }
}

// ============================================================
// 处理文本内容
// ============================================================
void Converter::handleText(const Token& tok) {
    if (curState_.skipGroup) return;
    // Feature 3：非组形式 htmlrtf 跳过
    if (htmlRtfInlineSkip_) return;

    const std::string& text = tok.text;
    if (text.empty()) return;

    switch (curState_.dest) {
        case Destination::ColorTable: {
            // 颜色表中 ; 分隔颜色条目
            // RTF 规范：第一个 ; 代表自动颜色（索引 0，已在初始化时添加）
            // 后续每个 ; 代表一个显式颜色条目
            for (char c : text) {
                if (c == ';') {
                    if (!colorTemp_.hasR && !colorTemp_.hasG && !colorTemp_.hasB) {
                        // 没有 \red\green\blue 前缀的 ';' 表示自动颜色条目
                        // 已经在初始化中添加了，但如果已经有非自动颜色则需要添加空条目
                        if (colorTable_.size() > 1) {
                            // 多余的自动条目，添加为自动
                            ColorEntry ce;
                            ce.isAuto = true;
                            colorTable_.push_back(ce);
                        }
                        // 否则跳过（已有 index 0 的自动颜色）
                    } else {
                        ColorEntry ce;
                        ce.isAuto = false;
                        ce.r = static_cast<uint8_t>(colorTemp_.r);
                        ce.g = static_cast<uint8_t>(colorTemp_.g);
                        ce.b = static_cast<uint8_t>(colorTemp_.b);
                        colorTable_.push_back(ce);
                    }
                    colorTemp_ = ColorParseTemp_{};
                }
            }
            return;
        }

        case Destination::FontEntry: {
            // 字体名称（以 ; 结尾）
            std::string fname;
            for (char c : text) {
                if (c == ';') {
                    if (currentFontEntry_.name.empty()) {
                        currentFontEntry_.name = fname;
                    } else if (currentFontEntry_.altName.empty()) {
                        currentFontEntry_.altName = fname;
                    }
                    if (currentFontIndex_ >= 0) {
                        fontTable_[currentFontIndex_] = currentFontEntry_;
                    }
                    fname.clear();
                } else {
                    fname += c;
                }
            }
            // 末尾没有 ; 的情况
            if (!fname.empty() && currentFontEntry_.name.empty()) {
                currentFontEntry_.name = fname;
            }
            return;
        }

        case Destination::Pict: {
            // 图片数据为十六进制字符串（空白分隔）
            for (char c : text) {
                if (std::isxdigit(c)) {
                    pict_.hexData += static_cast<char>(std::tolower(c));
                }
            }
            return;
        }

        case Destination::HtmlTag:
            if (fromHtml_ && collectingHtmlTag_) {
                htmlTagBuffer_ += text;
            }
            return;

        case Destination::FieldInst:
            field_.instText += text;
            return;

        case Destination::ListText:
            listTextBuffer_ += text;
            return;

        case Destination::StyleEntry:
            // Feature 2：收集样式名称（以 ; 结尾）
            {
                for (char c : text) {
                    if (c == ';') {
                        // 样式名称结束
                        currentStyle_.name = styleNameBuffer_;
                        styleNameBuffer_.clear();
                    } else {
                        styleNameBuffer_ += c;
                    }
                }
            }
            return;

        case Destination::BookmarkStart:
            // Feature 20：收集书签名称
            if (!fromHtml_) {
                for (char c : text) {
                    if (c != ';') bookmarkNameBuffer_ += c;
                }
            }
            return;

        case Destination::BookmarkEnd:
            return;

        case Destination::Footnote:
        case Destination::Endnote:
            // Feature 10：收集脚注文本
            if (collectingFootnote_) {
                footnoteBuffer_ += htmlEscape(text);
            }
            return;

        case Destination::Skip:
        case Destination::Info:
        case Destination::HtmlRtf:
        case Destination::MhtmlTag:
        case Destination::StyleSheet:
            return;

        default: {
            // 正常文本输出
            int cp = curState_.cs.fontCodepage > 0 ?
                     curState_.cs.fontCodepage : docCodepage_;
            std::string output;
            output.reserve(text.size());
            for (unsigned char c : text) {
                if (unicodeSkipCount_ > 0) {
                    --unicodeSkipCount_;
                    continue;
                }
                if (c < 0x80) {
                    output += static_cast<char>(c);
                } else {
                    output += cpToUtf8(c, cp);
                }
            }
            if (!output.empty()) {
                writeToCurrentOutput(htmlEscape(output));
            }
            return;
        }
    }
}

// ============================================================
// 写入到当前输出目标
// ============================================================
void Converter::writeToCurrentOutput(const std::string& s) {
    if (s.empty() || curState_.skipGroup) return;
    // Feature 3：非组形式 htmlrtf 跳过
    if (htmlRtfInlineSkip_) return;
    // Feature 10：脚注内容写入脚注缓冲
    if (collectingFootnote_ &&
        (curState_.dest == Destination::Footnote ||
         curState_.dest == Destination::Endnote)) {
        footnoteBuffer_ += s;
        return;
    }

    // HtmlTag 目标
    if (curState_.dest == Destination::HtmlTag) {
        if (fromHtml_ && collectingHtmlTag_) {
            htmlTagBuffer_ += s;
        }
        return;
    }

    // FieldInst 目标
    if (curState_.dest == Destination::FieldInst) {
        field_.instText += s;
        return;
    }

    // fromhtml 模式：Normal 和 FieldResult 目标直接输出
    if (fromHtml_) {
        if (curState_.dest == Destination::Normal ||
            curState_.dest == Destination::FieldResult) {
            html_ << s;
        }
        return;
    }

    // 普通模式：
    // 在表格单元格中
    if (curState_.ps.inTable && table_.active) {
        // 在单元格中打开 span（如果需要）
        std::string newCss = buildSpanCss();
        if (!newCss.empty()) {
            if (newCss != currentSpanCss_ || !spanOpen_) {
                if (spanOpen_) {
                    table_.currentCellContent += "</span>";
                    spanOpen_ = false;
                }
                table_.currentCellContent += "<span style=\"" + newCss + "\">";
                currentSpanCss_ = newCss;
                spanOpen_ = true;
            }
        } else if (spanOpen_) {
            table_.currentCellContent += "</span>";
            spanOpen_ = false;
            currentSpanCss_.clear();
        }
        table_.currentCellContent += s;
        return;
    }

    // 普通段落：先确保段落已打开，再打开 span
    if (!paraOpen_) {
        openParagraph();
    }

    // 打开或更新 span
    std::string newCss = buildSpanCss();
    if (!newCss.empty()) {
        if (newCss != currentSpanCss_ || !spanOpen_) {
            if (spanOpen_) {
                html_ << "</span>";
                spanOpen_ = false;
            }
            html_ << "<span style=\"" << newCss << "\">";
            currentSpanCss_ = newCss;
            spanOpen_ = true;
        }
    } else if (spanOpen_) {
        html_ << "</span>";
        spanOpen_ = false;
        currentSpanCss_.clear();
    }

    html_ << s;
}

// ============================================================
// 确保 span 已关闭
// ============================================================
void Converter::ensureSpanClosed() {
    if (!spanOpen_) return;

    if (curState_.ps.inTable && table_.active) {
        table_.currentCellContent += "</span>";
    } else if (paraOpen_) {
        html_ << "</span>";
    }
    spanOpen_ = false;
    currentSpanCss_.clear();
}

// ============================================================
// 确保 span 已按当前格式打开
// ============================================================
void Converter::ensureSpanOpen() {
    if (fromHtml_) return; // fromhtml 模式不管理 span

    std::string newCss = buildSpanCss();
    if (newCss == currentSpanCss_ && spanOpen_) return;

    if (spanOpen_) ensureSpanClosed();

    if (!newCss.empty()) {
        std::string tag = "<span style=\"" + newCss + "\">";
        if (curState_.ps.inTable && table_.active) {
            table_.currentCellContent += tag;
        } else {
            html_ << tag;
        }
        currentSpanCss_ = newCss;
        spanOpen_ = true;
    }
}

// ============================================================
// 构建字符格式 CSS
// ============================================================
std::string Converter::buildSpanCss() const {
    const CharState& cs = curState_.cs;
    std::ostringstream css;

    if (cs.hidden) {
        css << "display:none;";
        return css.str();
    }
    if (cs.bold)          css << "font-weight:bold;";
    if (cs.italic)        css << "font-style:italic;";

    if (cs.underline && cs.strikethrough)
        css << "text-decoration:underline line-through;";
    else if (cs.underline)
        css << "text-decoration:underline;";
    else if (cs.strikethrough)
        css << "text-decoration:line-through;";

    if (cs.superscript)   css << "vertical-align:super;font-size:smaller;";
    if (cs.subscript)     css << "vertical-align:sub;font-size:smaller;";

    double ptSize = cs.fontSize / 2.0;
    if (std::abs(ptSize - opts_.defaultFontSizePt) > 0.1)
        css << "font-size:" << ptSize << "pt;";

    if (cs.fgColorIndex > 0 &&
        cs.fgColorIndex < (int)colorTable_.size() &&
        !colorTable_[cs.fgColorIndex].isAuto)
        css << "color:" << colorToHex(cs.fgColorIndex) << ";";

    if (cs.highlightIndex > 0) {
        css << "background-color:" << highlightToColor(cs.highlightIndex) << ";";
    } else if (cs.bgColorIndex > 0 &&
               cs.bgColorIndex < (int)colorTable_.size() &&
               !colorTable_[cs.bgColorIndex].isAuto) {
        css << "background-color:" << colorToHex(cs.bgColorIndex) << ";";
    }

    std::string fontName = getFontName(cs.fontIndex);
    if (!fontName.empty()) {
        std::string fallback = "sans-serif";
        auto it = fontTable_.find(cs.fontIndex);
        if (it != fontTable_.end()) {
            switch (it->second.family) {
                case 1: fallback = "serif"; break;
                case 3: fallback = "monospace"; break;
                default: break;
            }
        }
        css << "font-family:\"" << fontName << "\"," << fallback << ";";
    }

    if (cs.smallCaps)     css << "font-variant:small-caps;";
    if (cs.allCaps)       css << "text-transform:uppercase;";

    if (cs.charSpacing != 0) {
        double spacingPt = cs.charSpacing / 20.0;
        css << "letter-spacing:" << spacingPt << "pt;";
    }

    // Feature 14：特殊文字效果
    if (cs.outline) {
        css << "-webkit-text-stroke:1px #000;color:transparent;";
    }
    if (cs.shadow) {
        css << "text-shadow:2px 2px 3px rgba(0,0,0,0.5);";
    }
    if (cs.emboss) {
        css << "text-shadow:1px 1px 0 #fff,-1px -1px 0 #888;";
    }
    if (cs.engrave) {
        css << "text-shadow:-1px -1px 0 #fff,1px 1px 0 #888;";
    }

    // Feature 9：字符方向
    if (cs.rtl) {
        css << "direction:rtl;unicode-bidi:bidi-override;";
    }

    return css.str();
}

// ============================================================
// 构建段落格式 CSS
// ============================================================
std::string Converter::buildParaCss() const {
    const ParaState& ps = curState_.ps;
    std::ostringstream css;

    switch (ps.align) {
        case 1: css << "text-align:center;"; break;
        case 2: css << "text-align:right;"; break;
        case 3: css << "text-align:justify;"; break;
        default: break;
    }
    if (ps.leftIndent != 0)
        css << "padding-left:" << twipsToPx(ps.leftIndent) << "px;";
    if (ps.rightIndent != 0)
        css << "padding-right:" << twipsToPx(ps.rightIndent) << "px;";
    if (ps.firstIndent != 0)
        css << "text-indent:" << twipsToPx(ps.firstIndent) << "px;";
    if (ps.spaceBefore != 0)
        css << "margin-top:" << twipsToPx(ps.spaceBefore) << "px;";
    if (ps.spaceAfter != 0)
        css << "margin-bottom:" << twipsToPx(ps.spaceAfter) << "px;";
    if (ps.lineSpacing != 0) {
        double absVal = std::abs((double)ps.lineSpacing);
        if (ps.lineSpacingRule == 0) {
            double lh = absVal / 240.0;
            if (lh > 0.1) css << "line-height:" << lh << ";";
        } else {
            double lhPt = absVal / 20.0;
            css << "line-height:" << lhPt << "pt;";
        }
    }

    // Feature 9：段落方向
    if (ps.rtlPar) {
        css << "direction:rtl;";
    }

    // Feature 11：分页
    if (ps.pageBreakBefore) {
        css << "page-break-before:always;";
    }

    // Feature 8：段落边框
    auto borderStyleCss = [](int style) -> std::string {
        switch (style) {
            case 1: return "thick";
            case 2: return "double";
            case 3: return "dotted";
            case 4: return "dashed";
            default: return "solid";
        }
    };
    if (ps.border.box) {
        double w = std::max(1.0, ps.border.width / 20.0);
        css << "border:" << w << "pt " << borderStyleCss(ps.border.style) << " ";
        if (ps.border.colorIndex > 0 && ps.border.colorIndex < (int)colorTable_.size()
            && !colorTable_[ps.border.colorIndex].isAuto)
            css << colorToHex(ps.border.colorIndex);
        else
            css << "#000";
        css << ";padding:4px;";
    } else {
        double w = std::max(1.0, ps.border.width / 20.0);
        std::string bcolor = "#000";
        if (ps.border.colorIndex > 0 && ps.border.colorIndex < (int)colorTable_.size()
            && !colorTable_[ps.border.colorIndex].isAuto)
            bcolor = colorToHex(ps.border.colorIndex);
        std::string bstyle = borderStyleCss(ps.border.style);
        if (ps.border.left)
            css << "border-left:" << w << "pt " << bstyle << " " << bcolor << ";";
        if (ps.border.right)
            css << "border-right:" << w << "pt " << bstyle << " " << bcolor << ";";
        if (ps.border.top)
            css << "border-top:" << w << "pt " << bstyle << " " << bcolor << ";";
        if (ps.border.bottom)
            css << "border-bottom:" << w << "pt " << bstyle << " " << bcolor << ";";
    }

    // Feature 15：段落底纹
    if (ps.shadingPct > 0) {
        // 混合前景色和背景色
        int fgR = 0, fgG = 0, fgB = 0;   // 默认黑色前景
        int bgR = 255, bgG = 255, bgB = 255; // 默认白色背景
        if (ps.shadingFgColor > 0 && ps.shadingFgColor < (int)colorTable_.size()
            && !colorTable_[ps.shadingFgColor].isAuto) {
            fgR = colorTable_[ps.shadingFgColor].r;
            fgG = colorTable_[ps.shadingFgColor].g;
            fgB = colorTable_[ps.shadingFgColor].b;
        }
        if (ps.shadingBgColor > 0 && ps.shadingBgColor < (int)colorTable_.size()
            && !colorTable_[ps.shadingBgColor].isAuto) {
            bgR = colorTable_[ps.shadingBgColor].r;
            bgG = colorTable_[ps.shadingBgColor].g;
            bgB = colorTable_[ps.shadingBgColor].b;
        }
        double pct = ps.shadingPct / 100.0;
        int mixR = static_cast<int>(fgR * pct + bgR * (1.0 - pct));
        int mixG = static_cast<int>(fgG * pct + bgG * (1.0 - pct));
        int mixB = static_cast<int>(fgB * pct + bgB * (1.0 - pct));
        char buf[24];
        std::snprintf(buf, sizeof(buf), "#%02x%02x%02x", mixR, mixG, mixB);
        css << "background-color:" << buf << ";";
    }

    return css.str();
}

// ============================================================
// 打开段落
// ============================================================
void Converter::openParagraph() {
    if (paraOpen_ || fromHtml_) return;

    // Feature 2：根据 outlinelevel 或样式名决定标签类型
    int outlineLevel = curState_.ps.outlineLevel;
    // 如果段落状态没有直接设置 outlinelevel，检查样式表
    if (outlineLevel < 0 && curState_.ps.styleIndex > 0) {
        auto it = styleTable_.find(curState_.ps.styleIndex);
        if (it != styleTable_.end()) {
            outlineLevel = it->second.outlineLevel;
        }
    }

    std::string tag = "p";
    if (outlineLevel >= 0 && outlineLevel <= 5) {
        tag = "h" + std::to_string(outlineLevel + 1);
    }

    std::string css = buildParaCss();
    if (!css.empty())
        html_ << "<" << tag << " style=\"" << css << "\">";
    else
        html_ << "<" << tag << ">";

    // 存储当前使用的标签（用于关闭）
    // 通过 paraOpen_ 保存标签类型到全局变量
    currentParaTag_ = tag;
    paraOpen_ = true;
    spanOpen_ = false;
    currentSpanCss_.clear();
}

// ============================================================
// 关闭段落
// ============================================================
void Converter::closeParagraph() {
    if (!paraOpen_ || fromHtml_) return;
    ensureSpanClosed();
    // Feature 2：使用对应的闭合标签
    html_ << "</" << currentParaTag_ << ">\n";
    paraOpen_ = false;
    currentParaTag_ = "p"; // 重置为默认
}

// ============================================================
// 发出段落结束（\par）
// ============================================================
void Converter::emitPar() {
    if (fromHtml_) return;

    if (curState_.ps.inTable && table_.active) {
        ensureSpanClosed();
        // 在表格单元格中，\par 是换行
        std::string& cc = table_.currentCellContent;
        // 避免在单元格末尾添加多余的 <br>
        if (!cc.empty()) {
            cc += "<br>";
        }
        return;
    }

    closeParagraph();
}

// ============================================================
// 单元格结束
// ============================================================
void Converter::emitCell() {
    if (!table_.active) return;
    ensureSpanClosed();

    std::string& content = table_.currentCellContent;
    // 移除末尾多余的 <br>
    while (content.size() >= 4 &&
           content.substr(content.size() - 4) == "<br>") {
        content.resize(content.size() - 4);
    }

    table_.cellContents.push_back(content);
    content.clear();
    table_.cellIndex++;
    spanOpen_ = false;
    currentSpanCss_.clear();
}

// ============================================================
// 行结束
// ============================================================
void Converter::emitRow() {
    if (!table_.active) return;

    // 处理未关闭的单元格
    if (!table_.currentCellContent.empty()) {
        emitCell();
    } else if (table_.cellIndex < (int)table_.cellDefs.size() &&
               table_.cellContents.size() < table_.cellDefs.size()) {
        // 空的最后一个单元格
        table_.cellContents.push_back("");
        table_.cellIndex++;
    }

    flushRow();
}

// ============================================================
// 开始表格
// ============================================================
void Converter::startTable() {
    if (table_.active) return;
    ensureSpanClosed();
    closeParagraph();
    html_ << "<table border=\"1\" "
          << "style=\"border-collapse:collapse;width:auto;\">\n";
    table_.active = true;
    table_.cellIndex = 0;
    paraOpen_ = false;
}

// ============================================================
// 结束表格
// ============================================================
void Converter::endTable() {
    if (!table_.active) return;

    // 输出剩余内容
    if (!table_.cellContents.empty() || !table_.currentCellContent.empty()) {
        if (!table_.currentCellContent.empty()) {
            emitCell();
        }
        if (!table_.cellContents.empty()) {
            flushRow();
        }
    }

    html_ << "</table>\n";
    table_.active = false;
    table_ = TableState{};
    paraOpen_ = false;
    spanOpen_ = false;
    currentSpanCss_.clear();
}

// ============================================================
// 输出当前行到 HTML
// ============================================================
void Converter::flushRow() {
    html_ << "<tr>";
    size_t n = table_.cellContents.size();
    for (size_t i = 0; i < n; ++i) {
        std::string tdStyle;
        if (i < table_.cellDefs.size()) {
            const TableCellDef& def = table_.cellDefs[i];
            std::ostringstream s;
            if (def.bgColorIndex > 0 &&
                def.bgColorIndex < (int)colorTable_.size() &&
                !colorTable_[def.bgColorIndex].isAuto)
                s << "background-color:" << colorToHex(def.bgColorIndex) << ";";
            switch (def.valign) {
                case 1: s << "vertical-align:middle;"; break;
                case 2: s << "vertical-align:bottom;"; break;
                default: s << "vertical-align:top;"; break;
            }
            // 设置宽度（如果有边界信息）
            if (def.rightBoundary > 0 && i > 0 && i - 1 < table_.cellDefs.size()) {
                int prev = (i > 0) ? table_.cellDefs[i-1].rightBoundary : 0;
                int width = def.rightBoundary - prev;
                if (width > 0) {
                    double widthPx = twipsToPx(width);
                    s << "width:" << (int)widthPx << "px;";
                }
            } else if (def.rightBoundary > 0 && i == 0) {
                double widthPx = twipsToPx(def.rightBoundary);
                s << "width:" << (int)widthPx << "px;";
            }
            tdStyle = s.str();
        }
        if (!tdStyle.empty())
            html_ << "<td style=\"" << tdStyle << "\">";
        else
            html_ << "<td>";
        html_ << table_.cellContents[i] << "</td>";
    }
    html_ << "</tr>\n";

    table_.cellContents.clear();
    table_.currentCellContent.clear();
    table_.cellIndex = 0;
    spanOpen_ = false;
    currentSpanCss_.clear();
}

// ============================================================
// 输出图片
// ============================================================
// 将十六进制字符串转二进制数组
std::vector<uint8_t> Converter::hexToBinary(const std::string& hex) {
    std::vector<uint8_t> result;
    result.reserve(hex.size() / 2);
    auto hexVal = [](char c) -> int {
        if (c >= '0' && c <= '9') return c - '0';
        if (c >= 'a' && c <= 'f') return c - 'a' + 10;
        if (c >= 'A' && c <= 'F') return c - 'A' + 10;
        return -1;
    };
    for (size_t i = 0; i + 1 < hex.size(); i += 2) {
        int hi = hexVal(hex[i]);
        int lo = hexVal(hex[i+1]);
        if (hi >= 0 && lo >= 0) {
            result.push_back(static_cast<uint8_t>(hi * 16 + lo));
        }
    }
    return result;
}

void Converter::emitPict() {
    if (pict_.format.empty()) return;

    // 将十六进制字符串转二进制
    std::vector<uint8_t> imgData = hexToBinary(pict_.hexData);

    // 合并二进制数据（\binN 方式）
    if (!pict_.binData.empty()) {
        imgData.insert(imgData.end(), pict_.binData.begin(), pict_.binData.end());
    }

    if (imgData.empty()) return;

    // 检查大小限制
    if (opts_.maxImageSizeBytes > 0 && imgData.size() > opts_.maxImageSizeBytes) {
        writeToCurrentOutput("<span>[图片过大，已跳过]</span>");
        return;
    }

    // 计算显示尺寸（twips 转像素）
    int dispW = pict_.widthGoal > 0 ? pict_.widthGoal : pict_.width;
    int dispH = pict_.heightGoal > 0 ? pict_.heightGoal : pict_.height;
    if (pict_.scaleX != 100 && dispW > 0) dispW = dispW * pict_.scaleX / 100;
    if (pict_.scaleY != 100 && dispH > 0) dispH = dispH * pict_.scaleY / 100;
    int wPx = dispW > 0 ? (int)twipsToPx(dispW) : 0;
    int hPx = dispH > 0 ? (int)twipsToPx(dispH) : 0;

    // Feature 6：WMF/EMF 处理
    if (pict_.format == "wmf" || pict_.format == "emf") {
        // 策略1：扫描数据中是否内嵌了 JPEG 或 PNG
        bool foundRaster = false;
        for (size_t i = 0; i + 3 < imgData.size(); ++i) {
            // 检查 JPEG 魔数：FF D8 FF
            if (imgData[i] == 0xFF && imgData[i+1] == 0xD8 && imgData[i+2] == 0xFF) {
                std::vector<uint8_t> jpegData(imgData.begin() + i, imgData.end());
                if (!opts_.embedImages) {
                    writeToCurrentOutput("<span>[图片]</span>");
                    return;
                }
                std::string b64 = toBase64(jpegData);
                std::ostringstream sizeAttr;
                if (wPx > 0) sizeAttr << " width=\"" << wPx << "\"";
                if (hPx > 0) sizeAttr << " height=\"" << hPx << "\"";
                std::string imgTag = "<img src=\"data:image/jpeg;base64," + b64 + "\""
                    + sizeAttr.str() + " alt=\"\" style=\"max-width:100%;\">";
                if (curState_.ps.inTable && table_.active)
                    table_.currentCellContent += imgTag;
                else {
                    if (!paraOpen_ && !fromHtml_) openParagraph();
                    html_ << imgTag;
                }
                foundRaster = true;
                break;
            }
            // 检查 PNG 魔数：89 50 4E 47
            if (imgData[i] == 0x89 && imgData[i+1] == 0x50 &&
                imgData[i+2] == 0x4E && imgData[i+3] == 0x47) {
                std::vector<uint8_t> pngData(imgData.begin() + i, imgData.end());
                if (!opts_.embedImages) {
                    writeToCurrentOutput("<span>[图片]</span>");
                    return;
                }
                std::string b64 = toBase64(pngData);
                std::ostringstream sizeAttr;
                if (wPx > 0) sizeAttr << " width=\"" << wPx << "\"";
                if (hPx > 0) sizeAttr << " height=\"" << hPx << "\"";
                std::string imgTag = "<img src=\"data:image/png;base64," + b64 + "\""
                    + sizeAttr.str() + " alt=\"\" style=\"max-width:100%;\">";
                if (curState_.ps.inTable && table_.active)
                    table_.currentCellContent += imgTag;
                else {
                    if (!paraOpen_ && !fromHtml_) openParagraph();
                    html_ << imgTag;
                }
                foundRaster = true;
                break;
            }
        }
        if (!foundRaster) {
            // 策略2：输出占位符
            std::ostringstream ph;
            ph << "<img";
            if (wPx > 0) ph << " width=\"" << wPx << "\"";
            if (hPx > 0) ph << " height=\"" << hPx << "\"";
            ph << " alt=\"[矢量图]\" style=\"";
            if (wPx > 0) ph << "width:" << wPx << "px;";
            if (hPx > 0) ph << "height:" << hPx << "px;";
            ph << "background:#f0f0f0;display:inline-block;"
               << "border:1px solid #ccc;vertical-align:middle;\">";
            if (curState_.ps.inTable && table_.active)
                table_.currentCellContent += ph.str();
            else {
                if (!paraOpen_ && !fromHtml_) openParagraph();
                html_ << ph.str();
            }
        }
        return;
    }

    // Feature 7：BMP/DIB 处理
    if (pict_.format == "bmp") {
        std::vector<uint8_t> pngData = dibToPng(imgData);
        if (!pngData.empty() && opts_.embedImages) {
            std::string b64 = toBase64(pngData);
            std::ostringstream sizeAttr;
            if (wPx > 0) sizeAttr << " width=\"" << wPx << "\"";
            if (hPx > 0) sizeAttr << " height=\"" << hPx << "\"";
            std::string imgTag = "<img src=\"data:image/png;base64," + b64 + "\""
                + sizeAttr.str() + " alt=\"\" style=\"max-width:100%;\">";
            if (curState_.ps.inTable && table_.active)
                table_.currentCellContent += imgTag;
            else {
                if (!paraOpen_ && !fromHtml_) openParagraph();
                html_ << imgTag;
            }
        } else if (!opts_.embedImages) {
            writeToCurrentOutput("<span>[图片]</span>");
        }
        return;
    }

    // PNG/JPEG 标准格式
    bool isSupportedFormat = (pict_.format == "png" || pict_.format == "jpeg");
    if (!isSupportedFormat) return;

    if (!opts_.embedImages) {
        writeToCurrentOutput("<span>[图片]</span>");
        return;
    }

    std::string b64 = toBase64(imgData);
    std::string mimeType = (pict_.format == "jpeg") ? "image/jpeg" : "image/png";

    std::ostringstream sizeAttr;
    if (wPx > 0) sizeAttr << " width=\"" << wPx << "\"";
    if (hPx > 0) sizeAttr << " height=\"" << hPx << "\"";

    std::string imgTag = "<img src=\"data:" + mimeType + ";base64," + b64 + "\""
                       + sizeAttr.str() + " alt=\"\" style=\"max-width:100%;\">";

    // 写入图片到正确位置
    if (curState_.ps.inTable && table_.active) {
        table_.currentCellContent += imgTag;
    } else {
        if (!paraOpen_ && !fromHtml_) openParagraph();
        html_ << imgTag;
    }
}

// ============================================================
// Feature 7：DIB/BMP 转 PNG（无压缩 PNG）
// ============================================================
std::vector<uint8_t> Converter::dibToPng(const std::vector<uint8_t>& dibData) {
    if (dibData.size() < 40) return {}; // BITMAPINFOHEADER 最小为 40 字节

    // 解析 BITMAPINFOHEADER
    auto readU32LE = [&](size_t off) -> uint32_t {
        if (off + 4 > dibData.size()) return 0;
        return (uint32_t)dibData[off] | ((uint32_t)dibData[off+1] << 8)
             | ((uint32_t)dibData[off+2] << 16) | ((uint32_t)dibData[off+3] << 24);
    };
    auto readU16LE = [&](size_t off) -> uint16_t {
        if (off + 2 > dibData.size()) return 0;
        return (uint16_t)dibData[off] | ((uint16_t)dibData[off+1] << 8);
    };

    uint32_t biSize       = readU32LE(0);
    int32_t  biWidth      = (int32_t)readU32LE(4);
    int32_t  biHeight     = (int32_t)readU32LE(8);
    uint16_t biBitCount   = readU16LE(14);
    uint32_t biCompression= readU32LE(16);

    // 只支持未压缩格式
    if (biCompression != 0) return {};
    if (biWidth <= 0) return {};
    // 高度为负表示自上到下，正表示自下到上
    bool topDown = (biHeight < 0);
    int height = topDown ? -((int32_t)biHeight) : (int32_t)biHeight;
    int width  = biWidth;

    // 颜色表数量
    size_t colorTableCount = 0;
    if (biBitCount <= 8) colorTableCount = (size_t(1) << biBitCount);
    size_t colorTableSize = colorTableCount * 4;
    size_t pixelDataOff = biSize + colorTableSize;

    if (pixelDataOff >= dibData.size()) return {};

    // 每行字节数（4字节对齐）
    size_t rowBytes = ((size_t)width * biBitCount + 31) / 32 * 4;

    // 生成 PNG（无压缩）
    // PNG 使用 zlib DEFLATE（BTYPE=00：存储块）包装扫描线
    // 每个扫描线前面有一个过滤器字节（0=None）

    // 构建 PNG 辅助函数
    auto crc32 = [](const uint8_t* data, size_t len) -> uint32_t {
        static uint32_t table[256];
        static bool init = false;
        if (!init) {
            for (uint32_t i = 0; i < 256; ++i) {
                uint32_t c = i;
                for (int k = 0; k < 8; ++k)
                    c = (c & 1) ? (0xEDB88320 ^ (c >> 1)) : (c >> 1);
                table[i] = c;
            }
            init = true;
        }
        uint32_t c = 0xFFFFFFFF;
        for (size_t i = 0; i < len; ++i) c = table[(c ^ data[i]) & 0xFF] ^ (c >> 8);
        return c ^ 0xFFFFFFFF;
    };

    auto adler32 = [](const uint8_t* data, size_t len) -> uint32_t {
        uint32_t a = 1, b = 0;
        for (size_t i = 0; i < len; ++i) {
            a = (a + data[i]) % 65521;
            b = (b + a) % 65521;
        }
        return (b << 16) | a;
    };

    auto writeU32BE = [](std::vector<uint8_t>& v, uint32_t val) {
        v.push_back((val >> 24) & 0xFF);
        v.push_back((val >> 16) & 0xFF);
        v.push_back((val >>  8) & 0xFF);
        v.push_back( val        & 0xFF);
    };

    auto writePngChunk = [&](std::vector<uint8_t>& png,
                              const char* type, const std::vector<uint8_t>& chunkData) {
        writeU32BE(png, (uint32_t)chunkData.size());
        png.push_back((uint8_t)type[0]); png.push_back((uint8_t)type[1]);
        png.push_back((uint8_t)type[2]); png.push_back((uint8_t)type[3]);
        png.insert(png.end(), chunkData.begin(), chunkData.end());
        // CRC 覆盖 type + data
        std::vector<uint8_t> crcInput;
        crcInput.push_back((uint8_t)type[0]); crcInput.push_back((uint8_t)type[1]);
        crcInput.push_back((uint8_t)type[2]); crcInput.push_back((uint8_t)type[3]);
        crcInput.insert(crcInput.end(), chunkData.begin(), chunkData.end());
        writeU32BE(png, crc32(crcInput.data(), crcInput.size()));
    };

    std::vector<uint8_t> png;

    // PNG 签名
    static const uint8_t pngSig[] = {0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A};
    png.insert(png.end(), pngSig, pngSig + 8);

    // IHDR 块（颜色类型 2=RGB, 6=RGBA）
    // 使用 RGB 输出（24位），palette 模式也转为 RGB
    {
        std::vector<uint8_t> ihdr(13);
        ihdr[0]  = (width >> 24) & 0xFF; ihdr[1]  = (width >> 16) & 0xFF;
        ihdr[2]  = (width >>  8) & 0xFF; ihdr[3]  =  width        & 0xFF;
        ihdr[4]  = (height >> 24) & 0xFF; ihdr[5]  = (height >> 16) & 0xFF;
        ihdr[6]  = (height >>  8) & 0xFF; ihdr[7]  =  height        & 0xFF;
        ihdr[8]  = 8;  // 位深度
        ihdr[9]  = 2;  // 颜色类型 RGB
        ihdr[10] = 0;  // 压缩方法 deflate
        ihdr[11] = 0;  // 过滤方法
        ihdr[12] = 0;  // 扫描方法（不隔行）
        writePngChunk(png, "IHDR", ihdr);
    }

    // 构建 RGB 扫描线数据（含过滤字节）
    std::vector<uint8_t> rawData;
    rawData.reserve((size_t)(width * 3 + 1) * height);

    auto getColor = [&](int x, int y) -> std::tuple<uint8_t, uint8_t, uint8_t> {
        // DIB 底部行优先（topDown=false 时 y=0 对应最底行）
        int srcRow = topDown ? y : (height - 1 - y);
        size_t rowOff = pixelDataOff + (size_t)srcRow * rowBytes;
        if (rowOff >= dibData.size())
            return std::make_tuple(uint8_t(0), uint8_t(0), uint8_t(0));

        if (biBitCount == 24) {
            size_t pOff = rowOff + (size_t)x * 3;
            if (pOff + 2 >= dibData.size())
                return std::make_tuple(uint8_t(0), uint8_t(0), uint8_t(0));
            uint8_t b = dibData[pOff], g = dibData[pOff+1], r = dibData[pOff+2];
            return std::make_tuple(r, g, b);
        } else if (biBitCount == 32) {
            size_t pOff = rowOff + (size_t)x * 4;
            if (pOff + 2 >= dibData.size())
                return std::make_tuple(uint8_t(0), uint8_t(0), uint8_t(0));
            uint8_t b = dibData[pOff], g = dibData[pOff+1], r = dibData[pOff+2];
            return std::make_tuple(r, g, b);
        } else if (biBitCount == 8 && colorTableCount > 0) {
            size_t pOff = rowOff + (size_t)x;
            if (pOff >= dibData.size())
                return std::make_tuple(uint8_t(0), uint8_t(0), uint8_t(0));
            uint8_t idx = dibData[pOff];
            size_t ctOff = biSize + (size_t)idx * 4;
            if (ctOff + 2 >= dibData.size())
                return std::make_tuple(uint8_t(0), uint8_t(0), uint8_t(0));
            uint8_t b = dibData[ctOff], g = dibData[ctOff+1], r = dibData[ctOff+2];
            return std::make_tuple(r, g, b);
        }
        return std::make_tuple(uint8_t(128), uint8_t(128), uint8_t(128)); // 不支持的位深度
    };

    for (int y = 0; y < height; ++y) {
        rawData.push_back(0); // 过滤器字节 0 = None
        for (int x = 0; x < width; ++x) {
            auto color = getColor(x, y);
            rawData.push_back(std::get<0>(color));
            rawData.push_back(std::get<1>(color));
            rawData.push_back(std::get<2>(color));
        }
    }

    // 使用 zlib 存储（非压缩）包装原始数据
    // zlib 格式：2字节头 + 一个或多个 DEFLATE 存储块 + 4字节 Adler-32
    std::vector<uint8_t> zdata;
    zdata.push_back(0x78); // CMF: deflate, window=32k
    zdata.push_back(0x01); // FLG: no dict, level 0 (fastest)

    // DEFLATE 存储块：每块最多 65535 字节
    size_t rawSize = rawData.size();
    size_t offset = 0;
    while (offset < rawSize) {
        size_t blockSize = std::min((size_t)65535, rawSize - offset);
        bool lastBlock = (offset + blockSize >= rawSize);
        // BFINAL + BTYPE=00（存储）
        zdata.push_back(lastBlock ? 0x01 : 0x00);
        // LEN (小端)
        zdata.push_back((uint8_t)(blockSize & 0xFF));
        zdata.push_back((uint8_t)((blockSize >> 8) & 0xFF));
        // NLEN（LEN 的补码）
        uint16_t nlen = (uint16_t)(~blockSize);
        zdata.push_back((uint8_t)(nlen & 0xFF));
        zdata.push_back((uint8_t)((nlen >> 8) & 0xFF));
        // 数据
        zdata.insert(zdata.end(), rawData.begin() + offset,
                                  rawData.begin() + offset + blockSize);
        offset += blockSize;
    }

    // Adler-32 校验（大端）
    uint32_t adler = adler32(rawData.data(), rawData.size());
    zdata.push_back((adler >> 24) & 0xFF);
    zdata.push_back((adler >> 16) & 0xFF);
    zdata.push_back((adler >>  8) & 0xFF);
    zdata.push_back( adler        & 0xFF);

    // IDAT 块
    writePngChunk(png, "IDAT", zdata);

    // IEND 块
    writePngChunk(png, "IEND", {});

    return png;
}

// ============================================================
// Base64 编码
// ============================================================
std::string Converter::toBase64(const std::vector<uint8_t>& data) {
    std::string result;
    result.reserve((data.size() + 2) / 3 * 4);

    size_t i = 0;
    for (; i + 2 < data.size(); i += 3) {
        uint32_t v = ((uint32_t)data[i] << 16) |
                     ((uint32_t)data[i+1] << 8) |
                     (uint32_t)data[i+2];
        result += kBase64Chars[(v >> 18) & 0x3F];
        result += kBase64Chars[(v >> 12) & 0x3F];
        result += kBase64Chars[(v >>  6) & 0x3F];
        result += kBase64Chars[(v      ) & 0x3F];
    }
    if (i + 1 == data.size()) {
        uint32_t v = (uint32_t)data[i] << 16;
        result += kBase64Chars[(v >> 18) & 0x3F];
        result += kBase64Chars[(v >> 12) & 0x3F];
        result += '=';
        result += '=';
    } else if (i + 2 == data.size()) {
        uint32_t v = ((uint32_t)data[i] << 16) | ((uint32_t)data[i+1] << 8);
        result += kBase64Chars[(v >> 18) & 0x3F];
        result += kBase64Chars[(v >> 12) & 0x3F];
        result += kBase64Chars[(v >>  6) & 0x3F];
        result += '=';
    }

    return result;
}

// ============================================================
// Feature 12：提取域指令内容，识别域类型
// ============================================================
std::string Converter::extractFieldContent(const std::string& instText, std::string& fieldType) {
    // 去掉首尾空白
    size_t s = instText.find_first_not_of(" \t\r\n");
    if (s == std::string::npos) { fieldType = ""; return ""; }
    std::string t = instText.substr(s);

    // 提取域类型（大写）
    size_t spacePos = t.find_first_of(" \t\r\n");
    std::string typeRaw = (spacePos != std::string::npos) ? t.substr(0, spacePos) : t;
    fieldType.resize(typeRaw.size());
    std::transform(typeRaw.begin(), typeRaw.end(), fieldType.begin(),
                   [](unsigned char c) { return (char)std::toupper(c); });

    std::string rest = (spacePos != std::string::npos) ? t.substr(spacePos) : "";
    // 去掉 rest 前的空白
    size_t rs = rest.find_first_not_of(" \t\r\n");
    if (rs != std::string::npos) rest = rest.substr(rs);

    if (fieldType == "HYPERLINK") {
        // 检查是否有 \l 本地锚点
        std::string url;
        bool hasLocal = false;
        // 先提取引号内的 URL
        if (!rest.empty() && rest[0] == '"') {
            size_t end = rest.find('"', 1);
            if (end != std::string::npos) url = rest.substr(1, end - 1);
        } else {
            size_t endPos = rest.find_first_of(" \t\r\n\\");
            url = (endPos != std::string::npos) ? rest.substr(0, endPos) : rest;
        }
        // 检查 \l 参数
        size_t lPos = rest.find("\\l");
        if (lPos != std::string::npos) {
            std::string afterL = rest.substr(lPos + 2);
            size_t ls2 = afterL.find_first_not_of(" \t");
            if (ls2 != std::string::npos) {
                afterL = afterL.substr(ls2);
                std::string anchor;
                if (!afterL.empty() && afterL[0] == '"') {
                    size_t ae = afterL.find('"', 1);
                    if (ae != std::string::npos) anchor = afterL.substr(1, ae - 1);
                } else {
                    size_t ae = afterL.find_first_of(" \t\r\n\\");
                    anchor = (ae != std::string::npos) ? afterL.substr(0, ae) : afterL;
                }
                if (!anchor.empty()) {
                    hasLocal = true;
                    if (url.empty()) url = "#" + anchor;
                    else url += "#" + anchor;
                }
            }
        }
        // 检查 \o 提示
        // （提示暂不在 extractFieldContent 中处理，需要在调用方处理）
        return url;
    } else if (fieldType == "REF") {
        // 提取书签名
        if (!rest.empty() && rest[0] == '"') {
            size_t end = rest.find('"', 1);
            if (end != std::string::npos) return rest.substr(1, end - 1);
        }
        size_t endPos = rest.find_first_of(" \t\r\n\\");
        return (endPos != std::string::npos) ? rest.substr(0, endPos) : rest;
    } else if (fieldType == "INCLUDEPICTURE") {
        if (!rest.empty() && rest[0] == '"') {
            size_t end = rest.find('"', 1);
            if (end != std::string::npos) return rest.substr(1, end - 1);
        }
        return rest;
    } else if (fieldType == "MAILTO") {
        if (!rest.empty() && rest[0] == '"') {
            size_t end = rest.find('"', 1);
            if (end != std::string::npos) return rest.substr(1, end - 1);
        }
        return rest;
    } else if (fieldType == "DATE" || fieldType == "TIME") {
        return "";
    } else if (fieldType == "PAGE" || fieldType == "NUMPAGES") {
        return "";
    }

    return rest;
}

// ============================================================
// 从 \fldinst 提取超链接 URL（保持向后兼容）
// ============================================================
std::string Converter::extractHyperlink(const std::string& instText) {
    std::string t = instText;
    // 去掉首尾空白
    auto s = t.find_first_not_of(" \t\r\n");
    if (s == std::string::npos) return "";
    t = t.substr(s);

    // 检查是否以 HYPERLINK 开头
    if (t.size() < 9) return "";
    std::string pfx = t.substr(0, 9);
    std::transform(pfx.begin(), pfx.end(), pfx.begin(),
                   [](unsigned char c) { return static_cast<char>(std::toupper(c)); });
    if (pfx != "HYPERLINK") return "";

    t = t.substr(9);
    s = t.find_first_not_of(" \t");
    if (s == std::string::npos) return "";
    t = t.substr(s);

    // 提取 URL
    std::string url;
    if (t[0] == '"') {
        auto e = t.find('"', 1);
        if (e != std::string::npos) url = t.substr(1, e - 1);
    } else {
        auto e = t.find_first_of(" \t\r\n\\");
        url = (e != std::string::npos) ? t.substr(0, e) : t;
    }

    return url;
}

// ============================================================
// HTML 转义
// ============================================================
std::string Converter::htmlEscape(const std::string& text) {
    std::string result;
    result.reserve(text.size() + 8);
    for (unsigned char c : text) {
        switch (c) {
            case '&': result += "&amp;"; break;
            case '<': result += "&lt;"; break;
            case '>': result += "&gt;"; break;
            case '"': result += "&quot;"; break;
            case '\r':
            case '\n': break; // RTF 文本中的换行已通过 \par 处理
            default:   result += (char)c; break;
        }
    }
    return result;
}

std::string Converter::htmlEscapeAttr(const std::string& text) {
    std::string result;
    for (unsigned char c : text) {
        switch (c) {
            case '&': result += "&amp;"; break;
            case '"': result += "&quot;"; break;
            case '<': result += "&lt;"; break;
            case '>': result += "&gt;"; break;
            default:  result += (char)c; break;
        }
    }
    return result;
}

// ============================================================
// 颜色转十六进制字符串
// ============================================================
std::string Converter::colorToHex(int index) const {
    if (index < 0 || index >= (int)colorTable_.size()) return "#000000";
    const ColorEntry& e = colorTable_[index];
    if (e.isAuto) return "inherit";
    char buf[8];
    std::snprintf(buf, sizeof(buf), "#%02x%02x%02x", e.r, e.g, e.b);
    return buf;
}

// ============================================================
// 高亮颜色索引转颜色字符串
// ============================================================
std::string Converter::highlightToColor(int index) const {
    static const char* colors[] = {
        "",          // 0
        "#FFFF00",   // 1 黄色
        "#00FF00",   // 2 亮绿
        "#00FFFF",   // 3 青色
        "#FF00FF",   // 4 洋红
        "#FF0000",   // 5 红色
        "#0000FF",   // 6 深蓝
        "#008000",   // 7 深绿
        "#008080",   // 8 深青
        "#800000",   // 9 深红
        "#800080",   // 10 深洋红
        "#808000",   // 11 深黄
        "#808080",   // 12 深灰
        "#C0C0C0",   // 13 浅灰
        "#000000",   // 14 黑色
        "#FFFFFF",   // 15 白色
    };
    if (index < 0 || index > 15) return "";
    return colors[index];
}

// ============================================================
// 获取字体名称
// ============================================================
std::string Converter::getFontName(int index) const {
    auto it = fontTable_.find(index);
    if (it == fontTable_.end()) return "";
    return it->second.name;
}

// ============================================================
// 更新当前字体代码页
// ============================================================
void Converter::updateFontCodepage(int fontIndex) {
    auto it = fontTable_.find(fontIndex);
    if (it != fontTable_.end()) {
        curState_.cs.fontCodepage = it->second.codepage;
        curState_.cs.fontCharset = it->second.charset;
    }
}

// ============================================================
// 列表管理
// ============================================================
void Converter::closeAllLists() {
    while (!listStack_.empty()) {
        const ListOutputState& ls = listStack_.back();
        html_ << (ls.isOrdered ? "</ol>\n" : "</ul>\n");
        listStack_.pop_back();
    }
}

void Converter::closeListsToLevel(int targetLevel) {
    while ((int)listStack_.size() > targetLevel + 1) {
        const ListOutputState& ls = listStack_.back();
        html_ << (ls.isOrdered ? "</ol>\n" : "</ul>\n");
        listStack_.pop_back();
    }
}

// ============================================================
// Feature 10：输出脚注/尾注区
// ============================================================
void Converter::emitFootnotes() {
    if (!footnotes_.empty()) {
        html_ << "<hr>\n<section class=\"footnotes\"><h4>脚注</h4>\n";
        for (const auto& fn : footnotes_) {
            html_ << fn;
        }
        html_ << "</section>\n";
    }
    if (!endnotes_.empty()) {
        html_ << "<hr>\n<section class=\"endnotes\"><h4>尾注</h4>\n";
        for (const auto& en : endnotes_) {
            html_ << en;
        }
        html_ << "</section>\n";
    }
}

} // namespace rtf2html
