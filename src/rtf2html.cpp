// rtf2html.cpp - 公共 API 实现
// 将 RTF 数据传递给 Converter 并返回结果

#include "rtf2html/rtf2html.h"
#include "converter.h"
#include <stdexcept>

namespace rtf2html {

// ============================================================
// 将 RTF 数据（指针+长度）转换为 HTML
// ============================================================
ConvertResult convert(const char* data, size_t size, const ConvertOptions& opts) {
    if (data == nullptr || size == 0) {
        ConvertResult result;
        result.success = false;
        result.errorMessage = "输入数据为空";
        return result;
    }

    try {
        Converter conv(data, size, opts);
        return conv.convert();
    } catch (const std::exception& e) {
        ConvertResult result;
        result.success = false;
        result.errorMessage = std::string("转换异常: ") + e.what();
        return result;
    } catch (...) {
        ConvertResult result;
        result.success = false;
        result.errorMessage = "未知异常";
        return result;
    }
}

// ============================================================
// 将 RTF 字符串转换为 HTML
// ============================================================
ConvertResult convert(const std::string& rtf, const ConvertOptions& opts) {
    return convert(rtf.data(), rtf.size(), opts);
}

} // namespace rtf2html
