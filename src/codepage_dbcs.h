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
