#pragma once
// tokenizer.h - RTF 词法分析器接口
// 将 RTF 字节流分解为 Token 序列

#include <string>
#include <vector>
#include <cstdint>
#include <cstddef>

namespace rtf2html {

// ============================================================
// Token 类型枚举
// ============================================================
enum class TokenType {
    GROUP_START,    // {
    GROUP_END,      // }
    CONTROL_WORD,   // \word[N]
    CONTROL_SYMBOL, // \x（单个非字母字符）
    HEX_CHAR,       // \'XX
    TEXT,           // 普通文本
    BINARY_DATA,    // \binN 后跟 N 字节二进制数据
    END_OF_INPUT,   // 输入结束
};

// ============================================================
// Token 结构
// ============================================================
struct Token {
    TokenType type = TokenType::END_OF_INPUT;

    // CONTROL_WORD: 控制字名称（小写）
    std::string name;

    // CONTROL_WORD: 是否有数字参数
    bool hasParam = false;

    // CONTROL_WORD: 数字参数值（可以为负）
    int param = 0;

    // CONTROL_SYMBOL: 控制符号字符
    char symbol = 0;

    // HEX_CHAR: 字节值
    uint8_t hexByte = 0;

    // TEXT: 文本内容
    std::string text;

    // BINARY_DATA: 二进制数据
    std::vector<uint8_t> binaryData;
};

// ============================================================
// RTF 词法分析器
// ============================================================
class Tokenizer {
public:
    /**
     * 构造函数
     * @param data  RTF 数据指针
     * @param size  数据长度
     */
    Tokenizer(const char* data, size_t size);

    /**
     * 读取下一个 Token
     * @return Token 对象
     */
    Token next();

    /**
     * 查看下一个 Token（不消费）
     * @return Token 对象
     */
    Token peek();

    /**
     * 是否已到达输入末尾
     */
    bool eof() const;

    /**
     * 获取当前位置（用于调试）
     */
    size_t position() const { return pos_; }

private:
    const char* data_;      // 数据指针
    size_t size_;           // 数据总长度
    size_t pos_;            // 当前读取位置

    bool hasPeeked_ = false;
    Token peekedToken_;

    // 读取下一个字符，返回 -1 表示 EOF
    int readChar();

    // 回退一个字符
    void unreadChar();

    // 查看当前字符（不消费）
    int peekChar() const;

    // 解析控制序列（以 \ 开头）
    Token parseControlSequence();

    // 解析文本 token
    Token parseText(char firstChar);

    // 解析二进制数据（\binN）
    Token parseBinaryData(int n);

    // 将十六进制字符转换为数值，失败返回 -1
    static int hexDigit(char c);
};

} // namespace rtf2html
