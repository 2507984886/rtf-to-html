// main.cpp - RTF 到 HTML 命令行工具
// 支持从文件或标准输入读取 RTF，输出 HTML

#include "rtf2html/rtf2html.h"
#include <iostream>
#include <fstream>
#include <sstream>
#include <string>
#include <cstring>
#include <cstdlib>

// Windows 需要额外的头文件来设置二进制模式
#if defined(_WIN32) || defined(_WIN64)
#  include <fcntl.h>
#  include <io.h>
#endif

// ============================================================
// 打印使用说明
// ============================================================
static void printUsage(const char* progName) {
    std::cerr
        << "用法: " << progName << " [选项] [输入文件.rtf]\n"
        << "\n"
        << "选项:\n"
        << "  -o <输出文件>    输出文件（默认：标准输出）\n"
        << "  -f               生成完整 HTML 页面（含 <html><head><body>）\n"
        << "  --no-images      跳过嵌入图片\n"
        << "  --font-size <N>  默认字体大小（磅，默认 12）\n"
        << "  --font <名称>    默认字体族（默认 \"Arial, sans-serif\"）\n"
        << "  --help           显示本帮助信息\n"
        << "\n"
        << "如果未指定输入文件，从标准输入读取。\n"
        << "\n"
        << "示例:\n"
        << "  " << progName << " email.rtf -o email.html\n"
        << "  " << progName << " -f email.rtf\n"
        << "  cat email.rtf | " << progName << " > email.html\n";
}

// ============================================================
// 从流中读取所有内容
// ============================================================
static std::string readStream(std::istream& in) {
    std::ostringstream oss;
    oss << in.rdbuf();
    return oss.str();
}

// ============================================================
// 从文件读取所有内容
// ============================================================
static bool readFile(const std::string& path, std::string& content, std::string& error) {
    std::ifstream f(path, std::ios::binary);
    if (!f.is_open()) {
        error = "无法打开文件: " + path;
        return false;
    }
    std::ostringstream oss;
    oss << f.rdbuf();
    if (f.fail() && !f.eof()) {
        error = "读取文件失败: " + path;
        return false;
    }
    content = oss.str();
    return true;
}

// ============================================================
// 主函数
// ============================================================
int main(int argc, char* argv[]) {
    // 解析命令行参数
    std::string inputFile;
    std::string outputFile;
    rtf2html::ConvertOptions opts;
    bool showHelp = false;

    for (int i = 1; i < argc; ++i) {
        const char* arg = argv[i];

        if (std::strcmp(arg, "--help") == 0 || std::strcmp(arg, "-h") == 0) {
            showHelp = true;
        } else if (std::strcmp(arg, "-f") == 0) {
            opts.generateFullPage = true;
        } else if (std::strcmp(arg, "--no-images") == 0) {
            opts.embedImages = false;
        } else if (std::strcmp(arg, "-o") == 0) {
            if (i + 1 < argc) {
                outputFile = argv[++i];
            } else {
                std::cerr << "错误: -o 选项需要指定输出文件\n";
                return 1;
            }
        } else if (std::strcmp(arg, "--font-size") == 0) {
            if (i + 1 < argc) {
                int sz = std::atoi(argv[++i]);
                if (sz > 0 && sz <= 200) {
                    opts.defaultFontSizePt = sz;
                } else {
                    std::cerr << "警告: 无效字体大小，使用默认值 12\n";
                }
            } else {
                std::cerr << "错误: --font-size 选项需要指定大小\n";
                return 1;
            }
        } else if (std::strcmp(arg, "--font") == 0) {
            if (i + 1 < argc) {
                opts.defaultFontFamily = argv[++i];
            } else {
                std::cerr << "错误: --font 选项需要指定字体名称\n";
                return 1;
            }
        } else if (arg[0] == '-') {
            std::cerr << "错误: 未知选项 " << arg << "\n";
            printUsage(argv[0]);
            return 1;
        } else {
            if (inputFile.empty()) {
                inputFile = arg;
            } else {
                std::cerr << "错误: 指定了多个输入文件\n";
                return 1;
            }
        }
    }

    if (showHelp) {
        printUsage(argv[0]);
        return 0;
    }

    // 读取输入
    std::string rtfContent;
    if (inputFile.empty()) {
        // 从标准输入读取
        // 在 Windows 上设置二进制模式
#if defined(_WIN32) || defined(_WIN64)
        _setmode(_fileno(stdin), _O_BINARY);
#endif
        rtfContent = readStream(std::cin);
        if (rtfContent.empty()) {
            std::cerr << "错误: 标准输入为空\n";
            return 1;
        }
    } else {
        std::string error;
        if (!readFile(inputFile, rtfContent, error)) {
            std::cerr << "错误: " << error << "\n";
            return 1;
        }
    }

    // 执行转换
    rtf2html::ConvertResult result = rtf2html::convert(rtfContent, opts);

    if (!result.success) {
        std::cerr << "转换警告: " << result.errorMessage << "\n";
        // 即使有警告也输出（可能有部分结果）
    }

    if (result.html.empty() && !result.success) {
        std::cerr << "转换失败，无输出\n";
        return 1;
    }

    // 写入输出
    if (outputFile.empty()) {
        // 写到标准输出
#if defined(_WIN32) || defined(_WIN64)
        _setmode(_fileno(stdout), _O_BINARY);
#endif
        std::cout << result.html;
        std::cout.flush();
    } else {
        std::ofstream outf(outputFile, std::ios::binary);
        if (!outf.is_open()) {
            std::cerr << "错误: 无法打开输出文件: " << outputFile << "\n";
            return 1;
        }
        outf << result.html;
        if (outf.fail()) {
            std::cerr << "错误: 写入输出文件失败: " << outputFile << "\n";
            return 1;
        }
        outf.close();

        if (!result.success) {
            std::cerr << "注意: 转换过程中遇到问题，输出可能不完整\n";
        } else {
            std::cerr << "成功: 已写入 " << outputFile
                      << " (" << result.html.size() << " 字节)\n";
        }
    }

    return result.success ? 0 : 2;
}
