#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# gen_dbcs_tables.py - 从 Unicode.org 映射文件生成 C++ DBCS 查找表
# 用法：python gen_dbcs_tables.py

import os

TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(os.path.dirname(TOOLS_DIR), "src")

# ============================================================
# 公共解析函数：读取映射文件，返回 {(lead, trail): unicode} 字典
# 只处理 DBCS 条目（源码 >= 0x8100，即高字节 >= 0x81）
# ============================================================
def parse_mapping(filepath):
    mapping = {}
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split()
            if len(parts) < 2:
                continue
            src_str = parts[0]
            uni_str = parts[1]
            if uni_str == "":
                continue
            try:
                src = int(src_str, 16)
                uni = int(uni_str, 16)
            except ValueError:
                continue
            # 只处理双字节条目
            if src < 0x8100:
                continue
            lead = (src >> 8) & 0xFF
            trail = src & 0xFF
            mapping[(lead, trail)] = uni
    return mapping

# ============================================================
# 生成 CP936 (GBK) 表
# lead: 0x81-0xFE (126行)
# trail: 0x40-0xFE 排除 0x7F (190列)
#   列索引: trail<=0x7E → trail-0x40; trail>=0x80 → trail-0x41
# ============================================================
def gen_gbk(mapping):
    ROWS = 126  # 0x81..0xFE
    COLS = 190  # 190个有效跟随字节
    table = [[0] * COLS for _ in range(ROWS)]

    for (lead, trail), uni in mapping.items():
        if lead < 0x81 or lead > 0xFE:
            continue
        if trail < 0x40 or trail > 0xFE or trail == 0x7F:
            continue
        row = lead - 0x81
        col = (trail - 0x40) if trail < 0x7F else (trail - 0x41)
        if uni <= 0xFFFF:
            table[row][col] = uni

    lines = []
    lines.append("// GBK (CP936) DBCS 双字节 → Unicode 查找表")
    lines.append("// 来源：https://www.unicode.org/Public/MAPPINGS/VENDORS/MICSFT/WINDOWS/CP936.TXT")
    lines.append("// 行索引 = lead - 0x81 (0..125)，列索引 = trail<0x7F ? trail-0x40 : trail-0x41 (0..189)")
    lines.append("static const uint16_t kGbkTable[126][190] = {")
    for row in range(ROWS):
        # 每行16个值一组
        chunks = []
        for c in range(0, COLS, 16):
            chunk = table[row][c:c+16]
            chunks.append(",".join(f"0x{v:04X}" for v in chunk))
        lines.append("    {" + ",".join(f"0x{v:04X}" for v in table[row]) + "},")
    lines.append("};")
    lines.append("")
    lines.append("// GBK 双字节转 Unicode")
    lines.append("static uint32_t gbkToUnicode(uint8_t lead, uint8_t trail) {")
    lines.append("    if (lead < 0x81 || lead > 0xFE) return 0xFFFD;")
    lines.append("    if (trail < 0x40 || trail > 0xFE || trail == 0x7F) return 0xFFFD;")
    lines.append("    int row = lead - 0x81;")
    lines.append("    int col = trail < 0x7F ? trail - 0x40 : trail - 0x41;")
    lines.append("    uint16_t u = kGbkTable[row][col];")
    lines.append("    return u ? u : 0xFFFD;")
    lines.append("}")
    return "\n".join(lines)

# ============================================================
# 生成 CP932 (Shift-JIS) 表
# lead: 0x81-0x9F (31个) + 0xE0-0xFC (29个) = 60行
# trail: 0x40-0x7E (63个) + 0x80-0xFC (125个) = 188列
# ============================================================
def gen_sjis(mapping):
    ROWS = 60
    COLS = 188
    table = [[0] * COLS for _ in range(ROWS)]

    def lead_to_row(lead):
        if 0x81 <= lead <= 0x9F:
            return lead - 0x81  # 0..30
        elif 0xE0 <= lead <= 0xFC:
            return (lead - 0xE0) + 31  # 31..59
        return -1

    def trail_to_col(trail):
        if 0x40 <= trail <= 0x7E:
            return trail - 0x40  # 0..62
        elif 0x80 <= trail <= 0xFC:
            return (trail - 0x80) + 63  # 63..187
        return -1

    for (lead, trail), uni in mapping.items():
        row = lead_to_row(lead)
        col = trail_to_col(trail)
        if row < 0 or col < 0:
            continue
        if uni <= 0xFFFF:
            table[row][col] = uni

    lines = []
    lines.append("// Shift-JIS (CP932) DBCS 双字节 → Unicode 查找表")
    lines.append("// 来源：https://www.unicode.org/Public/MAPPINGS/VENDORS/MICSFT/WINDOWS/CP932.TXT")
    lines.append("// 行: 0x81-0x9F→0..30, 0xE0-0xFC→31..59")
    lines.append("// 列: 0x40-0x7E→0..62, 0x80-0xFC→63..187")
    lines.append("static const uint16_t kSjisTable[60][188] = {")
    for row in range(ROWS):
        lines.append("    {" + ",".join(f"0x{v:04X}" for v in table[row]) + "},")
    lines.append("};")
    lines.append("")
    lines.append("// Shift-JIS 双字节转 Unicode")
    lines.append("static uint32_t sjisToUnicode(uint8_t lead, uint8_t trail) {")
    lines.append("    int row = -1;")
    lines.append("    if (lead >= 0x81 && lead <= 0x9F) row = lead - 0x81;")
    lines.append("    else if (lead >= 0xE0 && lead <= 0xFC) row = (lead - 0xE0) + 31;")
    lines.append("    if (row < 0) return 0xFFFD;")
    lines.append("    int col = -1;")
    lines.append("    if (trail >= 0x40 && trail <= 0x7E) col = trail - 0x40;")
    lines.append("    else if (trail >= 0x80 && trail <= 0xFC) col = (trail - 0x80) + 63;")
    lines.append("    if (col < 0) return 0xFFFD;")
    lines.append("    uint16_t u = kSjisTable[row][col];")
    lines.append("    return u ? u : 0xFFFD;")
    lines.append("}")
    return "\n".join(lines)

# ============================================================
# 生成 CP949 (韩语) 表
# lead: 0x81-0xFE (126行)
# trail: 0x41-0xFE (190列)
# ============================================================
def gen_korean(mapping):
    ROWS = 126
    COLS = 190  # 0x41..0xFE = 190个
    table = [[0] * COLS for _ in range(ROWS)]

    for (lead, trail), uni in mapping.items():
        if lead < 0x81 or lead > 0xFE:
            continue
        if trail < 0x41 or trail > 0xFE:
            continue
        row = lead - 0x81
        col = trail - 0x41
        if uni <= 0xFFFF:
            table[row][col] = uni

    lines = []
    lines.append("// CP949 (韩语) DBCS 双字节 → Unicode 查找表")
    lines.append("// 来源：https://www.unicode.org/Public/MAPPINGS/VENDORS/MICSFT/WINDOWS/CP949.TXT")
    lines.append("// 行索引 = lead - 0x81 (0..125)，列索引 = trail - 0x41 (0..189)")
    lines.append("static const uint16_t kKoreanTable[126][190] = {")
    for row in range(ROWS):
        lines.append("    {" + ",".join(f"0x{v:04X}" for v in table[row]) + "},")
    lines.append("};")
    lines.append("")
    lines.append("// CP949 双字节转 Unicode")
    lines.append("static uint32_t koreanToUnicode(uint8_t lead, uint8_t trail) {")
    lines.append("    if (lead < 0x81 || lead > 0xFE) return 0xFFFD;")
    lines.append("    if (trail < 0x41 || trail > 0xFE) return 0xFFFD;")
    lines.append("    int row = lead - 0x81;")
    lines.append("    int col = trail - 0x41;")
    lines.append("    uint16_t u = kKoreanTable[row][col];")
    lines.append("    return u ? u : 0xFFFD;")
    lines.append("}")
    return "\n".join(lines)

# ============================================================
# 生成 CP950 (Big5) 表
# lead: 0x81-0xFE (126行)
# trail: 0x40-0x7E (63个) + 0xA1-0xFE (94个) = 157列
# ============================================================
def gen_big5(mapping):
    ROWS = 126
    COLS = 157  # 63 + 94
    table = [[0] * COLS for _ in range(ROWS)]

    def trail_to_col(trail):
        if 0x40 <= trail <= 0x7E:
            return trail - 0x40  # 0..62
        elif 0xA1 <= trail <= 0xFE:
            return (trail - 0xA1) + 63  # 63..156
        return -1

    for (lead, trail), uni in mapping.items():
        if lead < 0x81 or lead > 0xFE:
            continue
        col = trail_to_col(trail)
        if col < 0:
            continue
        row = lead - 0x81
        if uni <= 0xFFFF:
            table[row][col] = uni

    lines = []
    lines.append("// Big5 (CP950) DBCS 双字节 → Unicode 查找表")
    lines.append("// 来源：https://www.unicode.org/Public/MAPPINGS/VENDORS/MICSFT/WINDOWS/CP950.TXT")
    lines.append("// 行索引 = lead - 0x81 (0..125)")
    lines.append("// 列索引: 0x40-0x7E→0..62, 0xA1-0xFE→63..156")
    lines.append("static const uint16_t kBig5Table[126][157] = {")
    for row in range(ROWS):
        lines.append("    {" + ",".join(f"0x{v:04X}" for v in table[row]) + "},")
    lines.append("};")
    lines.append("")
    lines.append("// Big5 双字节转 Unicode")
    lines.append("static uint32_t big5ToUnicode(uint8_t lead, uint8_t trail) {")
    lines.append("    if (lead < 0x81 || lead > 0xFE) return 0xFFFD;")
    lines.append("    int col = -1;")
    lines.append("    if (trail >= 0x40 && trail <= 0x7E) col = trail - 0x40;")
    lines.append("    else if (trail >= 0xA1 && trail <= 0xFE) col = (trail - 0xA1) + 63;")
    lines.append("    if (col < 0) return 0xFFFD;")
    lines.append("    int row = lead - 0x81;")
    lines.append("    uint16_t u = kBig5Table[row][col];")
    lines.append("    return u ? u : 0xFFFD;")
    lines.append("}")
    return "\n".join(lines)

# ============================================================
# 主程序
# ============================================================
def main():
    print("解析 CP936 (GBK)...")
    m936 = parse_mapping(os.path.join(TOOLS_DIR, "CP936.TXT"))
    print(f"  DBCS条目数: {len(m936)}")

    print("解析 CP932 (Shift-JIS)...")
    m932 = parse_mapping(os.path.join(TOOLS_DIR, "CP932.TXT"))
    print(f"  DBCS条目数: {len(m932)}")

    print("解析 CP949 (韩语)...")
    m949 = parse_mapping(os.path.join(TOOLS_DIR, "CP949.TXT"))
    print(f"  DBCS条目数: {len(m949)}")

    print("解析 CP950 (Big5)...")
    m950 = parse_mapping(os.path.join(TOOLS_DIR, "CP950.TXT"))
    print(f"  DBCS条目数: {len(m950)}")

    print("生成 C++ 表...")
    gbk_code  = gen_gbk(m936)
    sjis_code = gen_sjis(m932)
    kor_code  = gen_korean(m949)
    big5_code = gen_big5(m950)

    # 合并输出到 codepage_dbcs.cpp / codepage_dbcs.h
    header_path = os.path.join(SRC_DIR, "codepage_dbcs.h")
    impl_path   = os.path.join(SRC_DIR, "codepage_dbcs.cpp")

    # 头文件：只声明查找函数
    header = """\
// codepage_dbcs.h - DBCS (CJK 多字节编码) 到 Unicode 的完整查找表接口
// 由 tools/gen_dbcs_tables.py 自动生成，请勿手动修改
// 数据来源：https://www.unicode.org/Public/MAPPINGS/VENDORS/MICSFT/WINDOWS/

#pragma once
#include <cstdint>
#include <string>

namespace rtf2html {

// 各编码的双字节对转 UTF-8
std::string dbcsPairToUtf8(uint8_t lead, uint8_t trail, int codepage);

} // namespace rtf2html
"""

    # 实现文件
    impl = f"""\
// codepage_dbcs.cpp - DBCS (CJK 多字节编码) 到 Unicode 的完整查找表实现
// 由 tools/gen_dbcs_tables.py 自动生成，请勿手动修改
// 数据来源：https://www.unicode.org/Public/MAPPINGS/VENDORS/MICSFT/WINDOWS/
//
// 包含四个代码页的完整 DBCS→Unicode 映射表：
//   CP936 (GBK/简体中文): 126×190 = 23,940 项
//   CP932 (Shift-JIS/日语): 60×188 = 11,280 项
//   CP949 (韩语): 126×190 = 23,940 项
//   CP950 (Big5/繁体中文): 126×157 = 19,782 项

#include "codepage_dbcs.h"
#include "codepage.h"
#include <cstdint>

namespace rtf2html {{

// ============================================================
{gbk_code}

// ============================================================
{sjis_code}

// ============================================================
{kor_code}

// ============================================================
{big5_code}

// ============================================================
// 双字节对转 UTF-8（公共接口）
// ============================================================
std::string dbcsPairToUtf8(uint8_t lead, uint8_t trail, int codepage) {{
    uint32_t cp;
    switch (codepage) {{
        case 932: cp = sjisToUnicode(lead, trail);   break;
        case 936: cp = gbkToUnicode(lead, trail);    break;
        case 949: cp = koreanToUnicode(lead, trail); break;
        case 950: cp = big5ToUnicode(lead, trail);   break;
        default:  return "?";
    }}
    return utf32ToUtf8(cp == 0 ? 0xFFFD : cp);
}}

}} // namespace rtf2html
"""

    with open(header_path, "w", encoding="utf-8") as f:
        f.write(header)
    print(f"已写入 {header_path}")

    with open(impl_path, "w", encoding="utf-8") as f:
        f.write(impl)
    sz = os.path.getsize(impl_path)
    print(f"已写入 {impl_path} ({sz:,} 字节, {sz//1024} KB)")

if __name__ == "__main__":
    main()
