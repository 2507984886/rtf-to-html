#pragma once
// codepage.h - 代码页到 UTF-8 转换接口
// 支持 Windows 代码页和 DOS 代码页

#include <string>
#include <cstdint>

namespace rtf2html {

/**
 * 将单字节字符按指定代码页转换为 UTF-8 字符串
 * @param byte      原始字节值
 * @param codepage  代码页编号（如 1252, 1250 等）
 * @return          UTF-8 编码的字符串
 */
std::string cpToUtf8(unsigned char byte, int codepage);

/**
 * 将 RTF \ansicpg 值转换为代码页编号
 * @param ansicpg  \ansicpg 控制字的参数值
 * @return         代码页编号
 */
int ansicpgToCodepage(int ansicpg);

/**
 * 将 UTF-32 码点转换为 UTF-8 字符串
 * @param codepoint  Unicode 码点
 * @return           UTF-8 编码的字符串
 */
std::string utf32ToUtf8(uint32_t codepoint);

/**
 * 将 RTF \fcharset 值转换为代码页编号
 * @param fcharset  \fcharset 控制字的参数值
 * @return          代码页编号
 */
int fcharsetToCodepage(int fcharset);

/**
 * 判断字节是否为 DBCS 前导字节（Feature 1）
 * @param byte      字节值
 * @param codepage  代码页（932=Shift-JIS, 936=GBK, 949=Korean, 950=Big5）
 * @return          true 表示是前导字节
 */
bool isDbcsLeadByte(uint8_t byte, int codepage);

/**
 * DBCS 双字节对转 UTF-8（Feature 1）
 * @param lead      前导字节
 * @param trail     跟随字节
 * @param codepage  代码页
 * @return          UTF-8 编码的字符串（失败返回替换字符）
 */
std::string dbcsPairToUtf8(uint8_t lead, uint8_t trail, int codepage);

} // namespace rtf2html
