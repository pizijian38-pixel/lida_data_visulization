# lida-dataviz：Gemini 网页版（Gemini Apps）专用 skill

这是 `.agents/skills/lida-dataviz`（Antigravity 版）按 Gemini Apps 的 skill 规则改写的版本。
把它上传到 gemini.google.com 后，你在对话里上传数据、说一句需求，Gemini 就会生成：

- **PNG**：静态图表图片
- **XLSX**：原生可编辑的 Excel 图表，可以复制到 PPT

## 与 Antigravity 版的区别

| 规则（Gemini Apps 帮助文档） | Antigravity 版 | 本版本的处理 |
|---|---|---|
| 只能上传纯文本、PDF 和图片，`.xlsx` 等二进制文件不支持 | `examples/` 里有 `example_chart.xlsx` | 删除；打包脚本会拒绝任何二进制文件 |
| 隐藏的二进制文件（`.pyc`、`.DS_Store`）会导致上传失败 | 本地跑过脚本就会产生 `__pycache__/` | 打包脚本自动排除 |
| 脚本不能联网 | 第 0 步会 `pip install` | 不安装任何东西：先用 `check_env()` 检测环境，缺 openpyxl 只出 PNG，缺 seaborn 只用 matplotlib |
| 名称全小写、用连字符分隔；`SKILL.md` 在根目录 | 符合 | 符合，打包脚本会再检查一遍 |
| Gemini 只看 name + description 决定是否使用；说明要简短（cheat sheet, not manual）、给出示例 | `SKILL.md` 有 120 多行，面向本地工作区 | 重写为精简版，带交付示例和常见错误清单；细节放在 `references/` |
| 没有本地工作区和终端 | 依赖终端路径和 4 个相互 import 的脚本 | 合并为单文件 `scripts/lida_viz.py`，既能命令行运行，也能在 Python 会话里 import；文件不在磁盘上时，Gemini 可以先把内容写出来再用 |
| 沙箱里可能没有中文字体 | 要求安装字体 | 不能安装：PNG 改用英文文字，XLSX 照常用中文（由你的 Office 渲染） |

## 上传步骤

1. **打包**。在仓库根目录运行（会做规则检查，排除 `__pycache__` 等文件）：
   ```bash
   python gemini-skill/build_zip.py
   ```
   生成两份产物，任选一份上传：
   - `gemini-skill/dist/lida-dataviz/`（文件夹）
   - `gemini-skill/dist/lida-dataviz.zip`
2. **打开 Skills 页面**。在电脑上打开 gemini.google.com，用**个人 Google 账号**登录（工作或学校账号暂不支持），并开启 Keep Activity。侧边栏点 **Customize**（找不到就先点 More ⋯），再点顶部的 **Skills**。
3. **上传**。点顶部的 **Upload**，选择上面的文件夹或 zip，点 **Open**，检查无误后点 **Create**。
4. **使用**。在对话里上传数据文件，然后：
   - 输入 `/`（或 `@`），选择 **lida-dataviz** 并写需求，例如「画出各大区每月销售额的趋势」；
   - 也可以直接说需求，Gemini 会根据 description 自动使用这个 skill。
5. **更新**。以后修改了 skill，重新运行第 1 步，然后在 Skills 页面点该 skill 的 More ⋯ → **Replace skill** 上传。
   Gemini 要求更新时把整个 skill 连同参考文件重新上传一次。

## 测试

```bash
python gemini-skill/tests/test_lida_viz.py
```

## 尚未验证的地方

以下几点 Google 的文档没有写明，只能在真实的 Gemini 里试过才知道：

- **skill 里的 `.py` 文件是否直接出现在代码执行沙箱的磁盘上。** 不在磁盘上时，SKILL.md 会让 Gemini 把 `lida_viz.py` 的内容写成文件再 import。
- **沙箱里装了哪些库**（pandas、matplotlib 应该有；openpyxl、seaborn 不确定），**以及有没有中文字体。** 第一次使用时，Gemini 会运行 `check_env()`，结果会显示出来。
- **生成的 `.xlsx` 能否作为文件下载给你。** 如果不能，Gemini 会给出图表数据表，你可以粘贴到 Excel 后自己插入图表。

第一次试用后，请把 `check_env()` 的输出和遇到的问题发给我，我再据此调整。
