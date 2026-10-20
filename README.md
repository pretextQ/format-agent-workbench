# Format Agent Workbench

把格式要求和原始文档放进同一个工作台，生成可继续编辑的 Word。

**开源，个人及非商业使用免费。**

本项目仅面向 Windows 本机使用，从源码运行，不提供 Windows 便携包。

![Format Agent Workbench 首页](docs/images/workbench-home.png)

Format Agent Workbench 面向需要反复整理 Word 格式的人：上传格式要求或排版模板，再上传待处理文档，工作台会识别标题层级、字体字号、段落间距与页面设置，并交付排版后的 DOCX、修订模式 DOCX 和修改对照报告。

> 完整工作流需要一个能够同时理解文字和图片的多模态模型。内置演示使用项目中的固定样例与规则，不调用外部模型，也不需要 API Key。

## 三分钟开始

### Windows：从源码启动

1. 安装 Python 3.10 或更高版本，克隆或下载本仓库源码。
2. 双击 `启动工作台.bat`。首次运行会创建 `.venv` 并联网安装依赖。
3. 浏览器打开后，先开启“载入演示任务”，不用填写 API Key 即可体验完整交互。
4. 处理自己的文档时，点击右上角“模型设置”，选择服务并粘贴自己的 API Key。

以后仍然双击 `启动工作台.bat`；使用时请保留启动窗口，需要停止时按 Ctrl+C。模型设置保存在项目目录的 `.env`，历史产物保存在 `out/history`。

要求：64 位 Windows 10/11。为获得最稳定的 DOC/DOCX 转换与预览效果，建议安装 Microsoft Word；WPS 也可用于部分格式转换。没有 Word/WPS 时，DOCX 的基础排版仍可运行，但部分预览、旧格式转换与域刷新会受限。

## 使用自己的文档

真实任务会调用你配置的模型。点击工作台右上角的“模型设置”，然后：

1. 选择 Kimi Code、Kimi API、智谱 GLM、阿里云百炼、Gemini、OpenAI，或其他 OpenAI 兼容接口。
2. 粘贴 API Key。预设会自动填写接口地址和一个支持图片输入的模型名称，仍可自行修改。
3. 可先运行“测试多模态连接”，再点击“保存到本机”。测试会发送项目自带的一张示例图，产生一次很小的模型调用；设置保存后立即生效，不必重启。

API Key 只会写入项目目录中的 `.env`，该文件已被 Git 忽略，也不会进入排版结果。高级用户也可以复制 `.env.example` 为 `.env` 后手动编辑。

Temperature 默认使用“自动兼容”，即请求中不发送这个字段，由服务商采用模型默认值。这能避开 Kimi Code、Kimi K3 和部分推理模型因固定温度或不接受该字段而产生的 400 错误；只有服务商明确要求时才需要选择固定数值。

![在工作台内配置多模态模型](docs/images/model-settings.png)

请选择能同时接收文本与图片、并能稳定输出 JSON 的多模态模型。文字规则提取和文档结构识别主要使用文本能力；开启“视觉自检”后，模型还会检查渲染后的页面，因此图片理解能力是完整工作流的必要条件。

默认不会把文档页面上传到公共图床。详见 [隐私说明](PRIVACY.md)。

## 工作流

![填写要求并上传文档](docs/images/workbench-input.png)

1. **给出格式标准**：直接描述规则，或上传一份已经排好的 Word 模板。
2. **上传原始文档**：推荐 DOCX；Windows 下也可处理常见的 DOC、WPS、ODT、RTF。
3. **开始排版**：工作台识别结构、写入样式、检查结果，并保留正文内容。
4. **下载交付物**：排版后 DOCX、修订模式 DOCX、Markdown 修改报告。

![下载排版结果](docs/images/workbench-result.png)

### 一次排版会得到什么

- 可继续编辑的 DOCX，不是扁平图片或 PDF。
- Word 修订模式版本，可在审阅视图中逐项查看格式改动。
- 修改对照报告，记录段落角色、样式与修改项目。
- 可选视觉自检，对渲染页面进行第二轮检查。

| 排版前 | 排版后 |
| --- | --- |
| ![排版前](docs/images/before.png) | ![排版后](docs/images/after.png) |

## 支持范围与边界

- 最稳定的输入与输出格式是 DOCX；输出统一为 DOCX。
- 旧版 DOC、WPS、ODT、RTF 的可用性取决于本机 Word、WPS 或 LibreOffice。
- 工作台处理排版与结构，不替你改写正文事实，也不会从模板中复制作者姓名、声明等示例内容。
- 高度复杂的公式、宏、嵌入对象或特殊域建议在交付前人工复核。
- 视觉自检需要文档渲染环境和支持图片输入的模型。

## 命令行用法

图形界面之外，也可以直接调用排版 Agent：

```powershell
python main.py --spec assets/spec.txt --target draft.docx --out output/formatted.docx --verify
```

查看全部参数：

```powershell
python main.py --help
```

## 开发与验证

```powershell
python -m pip install -r requirements.txt
python -m pip install pytest
pytest tests/ -q
```

本仓库仅维护本地源码运行方式，不构建或发布 Windows EXE。历史版本记录见 [CHANGELOG.md](CHANGELOG.md)。

发现安全问题时，请不要公开提交包含真实文档、API Key 或个人信息的 Issue，处理方式见 [SECURITY.md](SECURITY.md)。

## 许可证

本项目采用 [PolyForm Noncommercial License 1.0.0](LICENSE)：源代码公开，个人学习、研究与其他非商业用途免费；商业使用需要另行取得授权。第三方依赖仍遵循其各自许可证。

