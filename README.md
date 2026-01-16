# remove-word-headers
# Word 页眉批量删除工具（支持图片页眉）

一个基于 Python 开发的 **Windows 批量处理工具**，  
用于 **递归删除指定文件夹及其所有子文件夹中 `.docx` 文件的页眉**，  
支持 **文字页眉 / 图片页眉 / 表格页眉**。

已打包为 `.exe`，**无需 Python 环境，双击即可使用**。

---

## ✨ 功能特性

- ✅ 批量处理 `.docx` 文件
- ✅ 递归处理所有子文件夹
- ✅ **彻底删除页眉（文字 / 图片 / 表格）**
- ✅ 自动跳过 Word 临时文件（`~$xxx.docx`）
- ✅ 控制台实时日志输出
- ✅ 支持直接运行 Python 脚本
- ✅ 提供 Windows 可执行文件（exe）

---

## 📁 使用方式（exe 版本，推荐）

### 使用步骤

1. 下载 `remove_word_headers.exe`
2. 将 exe **放入需要处理的“总文件夹”**
3. 双击运行
4. 程序会自动：
   - 扫描当前文件夹
   - 递归处理所有子文件夹中的 `.docx`
   - 删除所有页眉
5. 处理完成后，按回车键退出

> 📌 **exe 放在哪，就处理哪个文件夹**

---

## 🖥 运行效果示例

```text
开始处理目录：D:\Documents\WordFiles
已处理：D:\Documents\WordFiles\a.docx
已处理：D:\Documents\WordFiles\sub\b.docx
全部处理完成 ✅
按回车键退出...