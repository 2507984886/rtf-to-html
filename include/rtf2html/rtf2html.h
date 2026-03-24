#pragma once
// rtf2html.h - RTF 转 HTML 转换器公共 API
// 支持 Microsoft RTF 规范及 Outlook 特定扩展

#include <string>
#include <cstddef>

namespace rtf2html {

// ============================================================
// 转换选项
// ============================================================
struct ConvertOptions {
    // 是否将图片内嵌为 data URI（base64）
    bool embedImages = true;

    // 是否生成完整 HTML 页面（含 <html><head><body>）
    bool generateFullPage = false;

    // 默认字体大小（磅）
    int defaultFontSizePt = 12;

    // 默认字体族
    std::string defaultFontFamily = "Arial, sans-serif";

    // 图片最大字节数（超出则跳过）0 = 不限
    size_t maxImageSizeBytes = 0;
};

// ============================================================
// 转换结果
// ============================================================
struct ConvertResult {
    // 输出的 HTML 字符串
    std::string html;

    // 是否成功
    bool success = false;

    // 错误信息（失败时有效）
    std::string errorMessage;
};

// ============================================================
// 主转换函数
// ============================================================

/**
 * 将 RTF 数据转换为 HTML
 * @param data  RTF 数据指针
 * @param size  数据长度
 * @param opts  转换选项
 * @return      转换结果
 */
ConvertResult convert(const char* data, size_t size, const ConvertOptions& opts = ConvertOptions{});

/**
 * 将 RTF 字符串转换为 HTML
 * @param rtf   RTF 字符串
 * @param opts  转换选项
 * @return      转换结果
 */
ConvertResult convert(const std::string& rtf, const ConvertOptions& opts = ConvertOptions{});

} // namespace rtf2html
