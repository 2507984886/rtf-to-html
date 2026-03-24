// tokenizer.cpp - RTF 词法分析器实现
// 将 RTF 字节流按规范分解为 Token 序列

#include "tokenizer.h"
#include <cctype>
#include <stdexcept>

namespace rtf2html {

// ============================================================
// 构造函数
// ============================================================
Tokenizer::Tokenizer(const char* data, size_t size)
    : data_(data), size_(size), pos_(0)
{}

// ============================================================
// 字符级读取操作
// ============================================================
int Tokenizer::readChar() {
    if (pos_ >= size_) return -1;
    return static_cast<unsigned char>(data_[pos_++]);
}

void Tokenizer::unreadChar() {
    if (pos_ > 0) --pos_;
}

int Tokenizer::peekChar() const {
    if (pos_ >= size_) return -1;
    return static_cast<unsigned char>(data_[pos_]);
}

bool Tokenizer::eof() const {
    return pos_ >= size_ && !hasPeeked_;
}

// ============================================================
// 十六进制字符转数值
// ============================================================
int Tokenizer::hexDigit(char c) {
    if (c >= '0' && c <= '9') return c - '0';
    if (c >= 'a' && c <= 'f') return c - 'a' + 10;
    if (c >= 'A' && c <= 'F') return c - 'A' + 10;
    return -1;
}

// ============================================================
// 查看下一个 Token（不消费）
// ============================================================
Token Tokenizer::peek() {
    if (!hasPeeked_) {
        peekedToken_ = next();
        hasPeeked_ = true;
    }
    return peekedToken_;
}

// ============================================================
// 读取下一个 Token
// ============================================================
Token Tokenizer::next() {
    // 如果有预读的 token，直接返回
    if (hasPeeked_) {
        hasPeeked_ = false;
        return peekedToken_;
    }

    int c = readChar();
    if (c == -1) {
        Token t;
        t.type = TokenType::END_OF_INPUT;
        return t;
    }

    switch (c) {
        case '{': {
            Token t;
            t.type = TokenType::GROUP_START;
            return t;
        }
        case '}': {
            Token t;
            t.type = TokenType::GROUP_END;
            return t;
        }
        case '\\': {
            return parseControlSequence();
        }
        case '\r':
        case '\n': {
            // RTF 规范：换行符在 RTF 流中被忽略（不是 \par）
            // 继续读取下一个 token
            return next();
        }
        default: {
            return parseText(static_cast<char>(c));
        }
    }
}

// ============================================================
// 解析控制序列（已消费了 \ 字符）
// ============================================================
Token Tokenizer::parseControlSequence() {
    int c = readChar();
    if (c == -1) {
        // 文件末尾的反斜杠，作为文本处理
        Token t;
        t.type = TokenType::TEXT;
        t.text = "\\";
        return t;
    }

    Token tok;

    // 检查是否是字母（控制字）
    if (std::isalpha(c)) {
        // 读取控制字名称
        std::string name;
        name += static_cast<char>(std::tolower(c));
        while (true) {
            int nc = readChar();
            if (nc == -1) break;
            if (std::isalpha(nc)) {
                name += static_cast<char>(std::tolower(nc));
            } else {
                unreadChar();
                break;
            }
        }

        tok.type = TokenType::CONTROL_WORD;
        tok.name = name;
        tok.hasParam = false;
        tok.param = 0;

        // 检查是否有数字参数（可以有负号）
        int nc = readChar();
        if (nc == -1) {
            return tok;
        }

        bool negative = false;
        if (nc == '-') {
            // 负号
            int nc2 = readChar();
            if (nc2 != -1 && std::isdigit(nc2)) {
                negative = true;
                nc = nc2;
            } else {
                // 不是负数参数，回退
                if (nc2 != -1) unreadChar();
                // nc（'-'）处理为：如果下一个字符是空格则消费它，否则回退
                if (nc == ' ') {
                    // 空格是分隔符，消费掉
                } else {
                    unreadChar(); // 回退 '-'
                }
                return tok;
            }
        }

        if (std::isdigit(nc)) {
            // 读取数字参数
            std::string numStr;
            numStr += static_cast<char>(nc);
            while (true) {
                int nc2 = readChar();
                if (nc2 != -1 && std::isdigit(nc2)) {
                    numStr += static_cast<char>(nc2);
                } else {
                    if (nc2 != -1) {
                        if (nc2 != ' ') {
                            // 非空格终止符，回退
                            unreadChar();
                        }
                        // 空格是参数终止符，消费掉（不回退）
                    }
                    break;
                }
            }
            tok.hasParam = true;
            tok.param = std::stoi(numStr);
            if (negative) tok.param = -tok.param;

            // 特殊处理 \binN：读取 N 字节二进制数据
            if (tok.name == "bin" && tok.param > 0) {
                return parseBinaryData(tok.param);
            }
        } else {
            // 无数字参数
            if (nc == ' ') {
                // 空格是控制字终止符，消费掉（不回退）
            } else {
                unreadChar();
            }
        }

        return tok;
    }

    // 控制符号（单个非字母字符）
    tok.type = TokenType::CONTROL_SYMBOL;
    tok.symbol = static_cast<char>(c);

    switch (c) {
        case '\'': {
            // 十六进制转义 \'XX
            int hi = readChar();
            int lo = readChar();
            if (hi != -1 && lo != -1) {
                int h = hexDigit(static_cast<char>(hi));
                int l = hexDigit(static_cast<char>(lo));
                if (h >= 0 && l >= 0) {
                    tok.type = TokenType::HEX_CHAR;
                    tok.hexByte = static_cast<uint8_t>(h * 16 + l);
                    return tok;
                }
                // 无效的十六进制，回退
                unreadChar();
                unreadChar();
            }
            // 解析失败，作为控制符号返回
            tok.type = TokenType::CONTROL_SYMBOL;
            tok.symbol = '\'';
            return tok;
        }
        case '*': {
            // \* 可选目标标记
            tok.type = TokenType::CONTROL_SYMBOL;
            tok.symbol = '*';
            return tok;
        }
        case '\\':
        case '{':
        case '}': {
            // 转义的字面字符
            tok.type = TokenType::CONTROL_SYMBOL;
            tok.symbol = static_cast<char>(c);
            return tok;
        }
        case '~': {
            // 不换行空格
            tok.type = TokenType::CONTROL_SYMBOL;
            tok.symbol = '~';
            return tok;
        }
        case '-': {
            // 可选连字符
            tok.type = TokenType::CONTROL_SYMBOL;
            tok.symbol = '-';
            return tok;
        }
        case '_': {
            // 不换行连字符
            tok.type = TokenType::CONTROL_SYMBOL;
            tok.symbol = '_';
            return tok;
        }
        case ':': {
            // 索引子条目
            tok.type = TokenType::CONTROL_SYMBOL;
            tok.symbol = ':';
            return tok;
        }
        case '|': {
            // 公式字符
            tok.type = TokenType::CONTROL_SYMBOL;
            tok.symbol = '|';
            return tok;
        }
        default:
            // 其他控制符号
            return tok;
    }
}

// ============================================================
// 解析文本 Token
// ============================================================
Token Tokenizer::parseText(char firstChar) {
    Token tok;
    tok.type = TokenType::TEXT;
    tok.text += firstChar;

    while (true) {
        int c = readChar();
        if (c == -1) break;

        // 遇到特殊字符，停止文本读取
        if (c == '{' || c == '}' || c == '\\') {
            unreadChar();
            break;
        }
        // 跳过 RTF 流中的换行（根据规范，RTF 中的换行不是内容）
        if (c == '\r' || c == '\n') {
            break;
        }

        tok.text += static_cast<char>(c);
    }

    return tok;
}

// ============================================================
// 解析二进制数据（\binN 后紧跟 N 字节）
// ============================================================
Token Tokenizer::parseBinaryData(int n) {
    Token tok;
    tok.type = TokenType::BINARY_DATA;
    tok.binaryData.reserve(static_cast<size_t>(n));

    for (int i = 0; i < n; ++i) {
        int c = readChar();
        if (c == -1) break;
        tok.binaryData.push_back(static_cast<uint8_t>(c));
    }

    return tok;
}

} // namespace rtf2html
