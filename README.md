# WeChat Article Word Archiver Skill

一个用于归档微信公众号文章的 Codex Skill：将多篇公众号推文分别制作成带图 Word 文档，为每篇文章建立独立文件夹，同时保留原始图片，并生成归档清单和 ZIP 压缩包。

## 主要功能

- 批量处理 `mp.weixin.qq.com` 公众号文章链接。
- 每篇文章生成一份独立的 `.docx` 文档。
- 将文章中的静态图片嵌入 Word，同时保存到对应的 `图片` 子目录。
- 保留文章标题、公众号名称、发布日期和原始网页链接。
- 自动过滤过小的装饰图片、跟踪图片及 GIF 动图。
- 生成 `归档清单.json` 和完整 ZIP 压缩包。
- 检查 Word 内嵌图片数与独立图片文件数是否一致。
- 在 Windows 上通过 Microsoft Word 和 Poppler 渲染全部页面，辅助检查空白页、乱码、图片断裂和排版问题。
- 直接抓取受阻时，可切换到已经授权的 Chrome / Computer Use 流程。

## 适用场景

- 学校、机构或企业微信公众号内容归档。
- 将多篇活动报道分别整理为带图 Word。
- 按“编号 + 标题”建立独立文件夹。
- 批量制作可交付、可打印、可离线保存的公众号文章资料包。

## 环境要求

- Python 3.10 或更高版本。
- Python 包：`python-docx`、`lxml`、`Pillow`。
- Windows 全页渲染检查需要：
  - Microsoft Word；
  - Poppler 的 `pdftoppm`。
- 在 Codex 桌面环境中，优先使用内置的文档运行时和依赖。

如需手动安装 Python 依赖：

```powershell
python -m pip install python-docx lxml Pillow
```

## 安装为个人 Codex Skill

```powershell
$codexHome = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $env:USERPROFILE '.codex' }

git clone https://github.com/baqid1314/wechat-article-word-archiver.git `
  (Join-Path $codexHome 'skills\wechat-article-word-archiver')
```

如果 Skill 没有立即出现在 Codex 中，请重新启动 Codex，然后使用：

```text
$wechat-article-word-archiver
```

## 调用示例

只需提供归档名称，以及每篇文章的编号、文件名和链接：

```text
$wechat-article-word-archiver

归档名称：学校活动推文归档

01-活动启动仪式
https://mp.weixin.qq.com/s/...

02-比赛获奖喜报
https://mp.weixin.qq.com/s/...
```

也可以直接说明要求：

```text
$wechat-article-word-archiver 将下面这些公众号文章分别制作成带图 Word，
每篇放在独立文件夹中，并生成 ZIP 压缩包。
```

详细的 JSON 输入格式见 [`references/input-format.md`](references/input-format.md)。

## 输出结构

```text
<归档名称>/
  01-活动启动仪式/
    01-活动启动仪式.docx
    图片/
      图片01.jpg
      图片02.jpg
  02-比赛获奖喜报/
    02-比赛获奖喜报.docx
    图片/
      图片01.jpg
  归档清单.json

<归档名称>.zip
```

Word 文档内嵌的图片数量应与对应 `图片` 文件夹中的图片数量一致。

## 直接运行归档脚本

准备好 UTF-8 JSON 输入文件后，可以直接调用：

```powershell
python .\scripts\archive_wechat_articles.py `
  --input .\articles.json `
  --output-parent .\outputs `
  --cache-dir .\work\wechat-cache `
  --zip
```

仅在明确需要替换同名归档时增加 `--overwrite`。

## 质量检查

完整流程会检查：

1. 请求的文章数量与生成的文章文件夹数量一致。
2. 每个文章文件夹恰好包含一份 Word 和一个 `图片` 子目录。
3. Word 包含正文段落，且内嵌图片数与独立图片文件数一致。
4. 每份 Word 均能成功渲染，且至少包含一页。
5. 联系表覆盖所有渲染页面，方便逐页检查排版。
6. ZIP 可以正常打开，并包含全部 Word、图片和归档清单。

## 注意事项

- 只归档来源页面中真实存在的文字和静态图片，不补写缺失内容。
- CAPTCHA、安全提示、参数错误、登录页或空页面不能当作文章正文。
- 不绕过微信公众号或浏览器的安全检查。
- GIF、视频、音频和小程序等动态内容无法完整等价地保存在静态 Word 中。
- 如果公众号页面无法直接读取，应提供已保存的 HTML、PDF、完整正文或连续截图。
- 输入编号不连续时保留原编号，不自动补号或重新编号。

## 仓库结构

```text
SKILL.md
agents/
  openai.yaml
references/
  input-format.md
scripts/
  archive_wechat_articles.py
  make_contact_sheets.py
  render_docx_windows.ps1
```

