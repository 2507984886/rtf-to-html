// converter.cpp - RTF 到 HTML 转换器核心实现
// 处理 RTF 状态机、字体/颜色表、表格、图片、列表、fromhtml 模式

#include "converter.h"
#include "codepage.h"
#include <algorithm>
#include <cctype>
#include <cmath>
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
            // 域结果结束：如果有超链接，将已输出的内容包裹
            // 简化处理：已在 FieldResult 目标的文本中直接输出
            field_.inResult = false;
            break;
        }

        case Destination::FieldInst:
            field_.hyperlink = extractHyperlink(field_.instText);
            field_.inInst = false;
            // 如果是超链接，立即输出 <a> 开始标签
            if (!field_.hyperlink.empty()) {
                std::string href = "<a href=\"" + htmlEscapeAttr(field_.hyperlink) + "\">";
                if (!fromHtml_) {
                    // 在普通模式下，先确保段落打开
                    if (!paraOpen_) openParagraph();
                    ensureSpanClosed();
                    html_ << href;
                    fieldAnchorOpen_ = true;
                }
            }
            break;

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
        curState_.dest = Destination::StyleSheet;
        curState_.skipGroup = true;
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
        if (fromHtml_ && curState_.optionalDest) {
            curState_.dest = Destination::HtmlRtf;
            curState_.skipGroup = true;
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
    if (name == "bkmkstart" || name == "bkmkend") {
        curState_.dest = Destination::Bookmark;
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
        case Destination::StyleSheet:
        case Destination::StyleEntry:
        case Destination::ListText:
            // 跳过字体/颜色在这些目标中的处理
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
    if (name == "page" || name == "pagebb") {
        ensureSpanClosed();
        closeParagraph();
        html_ << "<div style=\"page-break-after:always\"></div>\n";
        return;
    }
    if (name == "sect" || name == "sectd") {
        return;
    }
    if (name == "s") {
        ps.styleIndex = param;
        return;
    }
    if (name == "nowidctlpar" || name == "widctlpar" ||
        name == "keepn" || name == "keep" ||
        name == "hyphpar" || name == "nohyph") {
        return;
    }
    (void)hasParam;
}

// ============================================================
// 表格控制字处理
// ============================================================
void Converter::handleTableControl(const std::string& name, bool /*hasParam*/, int param) {
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
}

// ============================================================
// 列表控制字处理
// ============================================================
void Converter::handleListControl(const std::string& name, bool /*hasParam*/, int param) {
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
            listOverrideTable_[param] = currentOverride_;
        }
        return;
    }
    if (name == "levelnfc" || name == "levelnfcn") {
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
    if (name == "listtemplateid" || name == "listid") return;
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
    if (name == "chpgn")      { writeToCurrentOutput("[PAGE]"); return; }
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

    if (unicodeSkipCount_ > 0) {
        --unicodeSkipCount_;
        return;
    }

    unsigned char byte = tok.hexByte;
    int cp = curState_.cs.fontCodepage > 0 ? curState_.cs.fontCodepage : docCodepage_;

    switch (curState_.dest) {
        case Destination::Pict:
            // 十六进制字符在 pict 中：已在 TEXT 中处理
            break;

        case Destination::FontEntry: {
            // 字体名称中的字符
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
            break;

        default: {
            std::string decoded = cpToUtf8(byte, cp);
            writeToCurrentOutput(htmlEscape(decoded));
            break;
        }
    }
}

// ============================================================
// 处理文本内容
// ============================================================
void Converter::handleText(const Token& tok) {
    if (curState_.skipGroup) return;

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

        case Destination::Skip:
        case Destination::Info:
        case Destination::HtmlRtf:
        case Destination::MhtmlTag:
        case Destination::StyleSheet:
        case Destination::StyleEntry:
        case Destination::Bookmark:
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
    return css.str();
}

// ============================================================
// 打开段落
// ============================================================
void Converter::openParagraph() {
    if (paraOpen_ || fromHtml_) return;
    std::string css = buildParaCss();
    if (!css.empty())
        html_ << "<p style=\"" << css << "\">";
    else
        html_ << "<p>";
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
    html_ << "</p>\n";
    paraOpen_ = false;
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
void Converter::emitPict() {
    if (pict_.format.empty()) return;

    bool isSupportedFormat = (pict_.format == "png" || pict_.format == "jpeg");
    if (!isSupportedFormat) return;

    // 将十六进制字符串转二进制
    std::vector<uint8_t> imgData;
    const std::string& hex = pict_.hexData;
    imgData.reserve(hex.size() / 2);
    for (size_t i = 0; i + 1 < hex.size(); i += 2) {
        auto hexVal = [](char c) -> int {
            if (c >= '0' && c <= '9') return c - '0';
            if (c >= 'a' && c <= 'f') return c - 'a' + 10;
            if (c >= 'A' && c <= 'F') return c - 'A' + 10;
            return 0;
        };
        imgData.push_back(static_cast<uint8_t>(hexVal(hex[i]) * 16 + hexVal(hex[i+1])));
    }

    // 合并二进制数据
    if (!pict_.binData.empty()) {
        imgData.insert(imgData.end(), pict_.binData.begin(), pict_.binData.end());
    }

    if (imgData.empty()) return;

    // 检查大小限制
    if (opts_.maxImageSizeBytes > 0 && imgData.size() > opts_.maxImageSizeBytes) {
        writeToCurrentOutput("<span>[图片过大，已跳过]</span>");
        return;
    }

    if (!opts_.embedImages) {
        writeToCurrentOutput("<span>[图片]</span>");
        return;
    }

    std::string b64 = toBase64(imgData);
    std::string mimeType = (pict_.format == "jpeg") ? "image/jpeg" : "image/png";

    // 计算显示尺寸（twips 转像素）
    std::ostringstream sizeAttr;
    int dispW = pict_.widthGoal > 0 ? pict_.widthGoal : pict_.width;
    int dispH = pict_.heightGoal > 0 ? pict_.heightGoal : pict_.height;
    if (pict_.scaleX != 100 && dispW > 0) dispW = dispW * pict_.scaleX / 100;
    if (pict_.scaleY != 100 && dispH > 0) dispH = dispH * pict_.scaleY / 100;
    if (dispW > 0) sizeAttr << " width=\"" << (int)twipsToPx(dispW) << "\"";
    if (dispH > 0) sizeAttr << " height=\"" << (int)twipsToPx(dispH) << "\"";

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
// 从 \fldinst 提取超链接 URL
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

} // namespace rtf2html
