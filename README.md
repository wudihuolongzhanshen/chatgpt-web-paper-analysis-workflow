# ChatGPT 网页版批量论文分析工作流（无需 API）

本项目用于在 Windows 上批量处理论文 PDF：每篇论文都在已登录的 ChatGPT 网页版中新建一个独立对话，只上传这一篇 PDF，并发送统一的分析提示词；回答完成后，扩展自动点击 ChatGPT 的“复制回答”按钮，再调用本机 Microsoft Word，以“保留源格式”粘贴并保存分析结果。

当前 Chrome 扩展版本：`1.1.1`。

## 1. 这个项目做什么

完整自动化流程如下：

1. 在本项目的扩展窗口中拖入或选择一个或多个 PDF。
2. 本地服务为每篇论文创建任务、输出目录和 `progress.csv` 记录。
3. 扩展按照任务顺序逐篇处理。
4. 每篇论文都回到 ChatGPT 空白首页并建立新对话。
5. 当前对话只上传这一篇论文的 PDF。
6. 扩展读取 `prompt/analysis_prompt.md` 的最新完整内容并发送。
7. 等待 ChatGPT 回答结束。
8. 通过 Chrome 可信鼠标事件点击当前回答底部的“复制回答”按钮。
9. 校验剪贴板内容确实属于当前回答；不匹配时最多自动重试三次。
10. 本地服务启动 Microsoft Word，执行“保留源格式”粘贴。
11. Word 保存后，再次检查文档正文是否与当前 ChatGPT 回答一致。
12. 验证通过后保存论文名 `.docx`，并把任务设为 `completed`。
13. 单篇失败时记录原因并继续下一篇，同一批次不会立即重复失败任务。

本项目不调用 OpenAI API，也不需要 API Key。论文分析本身使用你已登录的 ChatGPT 网页会话。

## 2. 需要准备什么

### 2.1 操作系统

当前实现面向 Windows，原因是最终 Word 文件通过 Windows 的 Microsoft Word COM 自动化接口生成。

建议：

- Windows 10 或 Windows 11；
- 使用具有本地文件读写权限的普通桌面账户；
- 项目路径尽量不要再嵌套到更深的目录，避免 Windows 路径过长。

### 2.2 Google Chrome

需要安装桌面版 Google Chrome。

不需要安装任何第三方“GPT 插件”，也不需要安装官方 Codex Chrome 插件。这个项目使用的是仓库自带的个人 Chrome 扩展：

```text
browser-extension/
```

扩展直接运行在 `chatgpt.com` 页面和本地任务窗口中。

### 2.3 ChatGPT 网页账号

需要：

- 能正常登录 `https://chatgpt.com/`；
- 当前账号能够在网页对话中上传 PDF；
- 当前账号能够生成完整论文分析回答。

本项目按你的 ChatGPT Business 网页账号设计。运行前请先在 Chrome 中手动登录，并确认可以正常新建对话、上传一个 PDF、发送消息和点击“复制回答”。

如果网页出现验证码、账号确认、安全提示或使用额度限制，需要人工处理；扩展不会尝试绕过这些限制。

### 2.4 Microsoft Word 桌面版

必须安装可正常启动的 Microsoft Word 桌面版。仅安装网页版 Word 不够。

可用 PowerShell 检查 Word 自动化接口：

```powershell
$word = New-Object -ComObject Word.Application
$word.Version
$word.Quit()
```

如果能输出 Word 版本号，说明接口可用。

运行批处理时，建议关闭正在编辑的任务结果 Word，避免文件被锁定而无法替换。

### 2.5 Python

建议安装 Python 3.10 或更高版本。在 PowerShell 中检查：

```powershell
py --version
```

如果系统找不到 `py`，请安装 Windows 版 Python，并在安装时启用 Python Launcher 或将 Python 加入 PATH。

### 2.6 快速启动
```powershell
git clone https://github.com/wudihuolongzhanshen/chatgpt-web-paper-analysis-workflow.git
```
```powershell
cd chatgpt-web-paper-analysis-workflow
```
```powershell
py -m pip install -r requirements.txt
```
```powershell
py workflow/local_server.py
```

## 3. 项目目录结构

关键目录和文件：

```text
项目根目录/
├─ browser-extension/          # 本项目自己的 Chrome 扩展
├─ output/                     # PDF 原文和 Word 结果
├─ prompt/
│  └─ analysis_prompt.md       # 所有论文共用的提示词
├─ workflow/
│  ├─ local_server.py          # 本地任务服务
│  ├─ workflow_common.py       # 路径、状态和 CSV 公共逻辑
│  ├─ check_results.py         # Word 本地检查
│  └─ get_next_task.py         # 可选的命令行任务领取工具
├─ progress.csv                # 任务列表和状态
├─ requirements.txt            # Python 依赖
└─ .workflow-token             # 本地扩展访问令牌，首次启动生成
```


典型输出结构：

```text
output/
└─ 001_论文标题开头/
   ├─ 完整论文标题.pdf
   └─ 完整论文标题.docx
```

为避免 Windows 路径过长，任务文件夹中的论文名部分会适度截短；文件夹内部的 PDF 和 Word 会尽量保留完整论文名。

## 4. 第一次安装

### 4.1 打开项目目录

先克隆公开仓库，再进入项目根目录：

```powershell
git clone https://github.com/wudihuolongzhanshen/chatgpt-web-paper-analysis-workflow.git
cd chatgpt-web-paper-analysis-workflow
```

如果没有安装 Git，也可以在 GitHub 仓库页面选择 **Code → Download ZIP**，解压后使用 PowerShell 进入解压目录。

后续命令都应在这个目录运行。

### 4.2 安装 Python 依赖

```powershell
py -m pip install -r requirements.txt
```

当前核心依赖是 `python-docx`，用于检查生成的 Word 是否可打开、大小是否合理，并读取文档正文。

如果出现权限问题，可尝试：

```powershell
py -m pip install --user -r requirements.txt
```

### 4.3 设置统一提示词

打开：

```text
prompt/analysis_prompt.md
```

把你希望每篇论文都执行的完整分析要求写入这个文件并保存为 UTF-8。

注意：

- 所有任务都读取同一个文件；
- 每次领取任务时读取最新内容，不使用输出目录中的旧副本；
- 不要要求 ChatGPT 生成或下载 Word；
- 应要求 ChatGPT 直接在网页回答中输出完整分析；
- 如果修改提示词，只影响修改后新发送的任务，不会改变已经发送的对话。

### 4.4 启动本地服务

```powershell
py workflow/local_server.py
```

正常启动后终端会显示：

- 本地地址：`http://127.0.0.1:8765`；
- 扩展令牌；
- 服务仅监听本机。

首次运行会在根目录生成：

```text
.workflow-token
```

这个文件里是一行随机令牌。扩展必须携带正确令牌才能读取或修改任务。

保持这个 PowerShell 窗口运行。关闭窗口或按 `Ctrl+C` 会停止本地服务，扩展随后会提示无法连接。

可在另一个 PowerShell 窗口检查服务：

```powershell
Invoke-RestMethod http://127.0.0.1:8765/api/health
```

正常结果应包含：

```json
{"ok": true, "service": "paper-workflow"}
```

## 5. 安装项目自带的 Chrome 扩展

### 5.1 加载未打包扩展

1. 在 Chrome 地址栏打开：

   ```text
   chrome://extensions/
   ```

2. 打开右上角“开发者模式”。
3. 点击“加载已解压的扩展程序”。
4. 选择项目中的：

   ```text
   browser-extension
   ```

5. 确认扩展显示为“本地论文逐篇分析”或“论文分析队列”。
6. 确认版本是 `1.1.1`。
7. 建议点击图钉，将扩展固定到 Chrome 工具栏。

选择的是整个 `browser-extension` 文件夹，不是其中的 `manifest.json`。

### 5.2 扩展权限说明

扩展会申请以下权限：

- `storage`：保存本地服务地址、令牌、选中任务和运行状态；
- `tabs`：找到并控制已登录的 ChatGPT 标签页；
- `scripting`：在 ChatGPT 页面注入自动化脚本；
- `debugger`：发送浏览器认可的可信鼠标事件，真正点击“复制回答”；
- `windows`：打开和聚焦独立的任务管理窗口；
- `system.display`：按显示器可用区域设置任务窗口高度；
- `clipboardRead`：确认复制后的剪贴板内容属于当前回答。

允许访问的地址只有：

- `https://chatgpt.com/*`；
- `http://127.0.0.1:8765/*`。

使用可信点击时，Chrome 可能短暂显示“扩展正在调试此浏览器”之类提示，这是 `debugger` 权限的正常表现。

### 5.3 更新扩展

项目代码修改后，Chrome 不会总是自动加载最新脚本。更新步骤：

1. 打开 `chrome://extensions/`；
2. 找到本扩展；
3. 点击“重新加载”；
4. 确认版本号；
5. 关闭旧任务窗口；
6. 再次点击扩展图标打开新窗口；
7. 必要时刷新 ChatGPT 页面。

如果新增了权限，Chrome 可能要求再次确认。

## 6. 第一次连接扩展和本地服务

1. 确认 `py workflow/local_server.py` 正在运行。
2. 在 Chrome 中登录 `https://chatgpt.com/`。
3. 点击工具栏中的扩展图标。
4. 扩展会打开一个独立任务窗口；再次点击图标会聚焦已有窗口，不会重复创建。
5. “本地服务地址”保持：

   ```text
   http://127.0.0.1:8765
   ```

6. 打开项目根目录的 `.workflow-token`，复制整行令牌。
7. 粘贴到扩展的“扩展令牌”。
8. 点击“刷新”。

如果连接正常，状态区会显示“已停止”，并显示当前任务列表。

## 7. 导入论文 PDF

### 7.1 拖拽导入

1. 在资源管理器中选择一个或多个 PDF。
2. 把文件拖入扩展任务窗口的任意位置。
3. 只有检测到文件拖入时，扩展才会显示全屏投放提示。
4. 松开鼠标后开始上传。
5. 状态区会显示当前文件和上传进度。
6. 上传完成后任务立即出现在列表中，默认状态为 `pending`。

### 7.2 文件选择器导入

点击“选择 PDF”，可一次选择多个 PDF。

### 7.3 导入规则

- 只接受 `.pdf`；
- 单个文件最大 200 MB；
- 服务检查文件头是否为 PDF；
- 同名论文已存在时拒绝覆盖；
- PDF 直接保存到对应的 `output` 任务目录；
- 导入时立即更新 `progress.csv`；
- 不创建 `papers/`；
- 上传中断时会清理临时目录，避免留下半个任务。

## 8. 开始批量处理

### 8.1 运行前检查

确认：

- 本地服务正在运行；
- Chrome 已登录 ChatGPT；
- ChatGPT 页面可以正常新建对话；
- `prompt/analysis_prompt.md` 已保存；
- Microsoft Word 桌面版可用；
- 没有打开将被替换的任务 Word；
- 扩展版本与项目版本一致。

### 8.2 选择任务

任务窗口提供：

- “选择未完成”：选择所有状态不是 `completed` 的任务；
- “选择全部”：选择列表中所有任务；
- 再次点击相应按钮可取消选择；
- “刷新”只更新列表，不会擅自增加选择。

点击“保存并开始”时，只会执行所选任务中状态为：

- `pending`；
- `needs_review`。

即使选中了 `completed`，也不会重复执行它。

### 8.3 启动

1. 先打开一个已登录的 ChatGPT 标签页，建议停留在空白首页。
2. 在扩展任务窗口勾选任务。
3. 点击“保存并开始”。
4. 扩展会锁定一个 ChatGPT 标签页作为受控标签页。
5. 其他 ChatGPT 标签页不会自动加入队列。

处理期间不要在受控标签页中手动发送消息、切换对话或上传其他文件。上传、发送和等待回答可以在后台进行，但浏览器出于剪贴板安全限制，复制回答时扩展会自动把受控 ChatGPT 标签页及其 Chrome 窗口切到前台；这是正常行为。

### 8.4 单篇任务内部过程

扩展会自动：

1. 回到 `https://chatgpt.com/` 空白首页；
2. 上传当前任务 PDF；
3. 等待页面显示附件上传完成；
4. 把共享提示词写入输入框，并保留换行；
5. 发送消息；
6. 最长等待约 45 分钟；
7. 等待回答停止且文本稳定；
8. 找到当前回答专用的 `copy-turn-action-button`；
9. 排除“复制表格”“复制代码”等内部按钮；
10. 通过可信鼠标事件点击；
11. 读取并核对剪贴板；
12. 启动 Word 原格式粘贴；
13. 验证 Word 正文；
14. 保存对话 URL 和任务状态；
15. 进入下一篇的新对话。

## 9. Word 输出规则

本项目不使用 HTML 转 Word，也没有 HTML 回退渠道。

唯一正式路径是：

```text
ChatGPT“复制回答” → 系统剪贴板 → Microsoft Word“保留源格式”粘贴 → .docx
```

生成后至少检查：

- Word 文件存在；
- 文件大小至少 20 KB；
- `python-docx` 可以打开；
- Word 正文与当前回答一致。

ChatGPT 网页中的视觉标题不一定成为 Word 内置的 Heading 样式，所以自动流程不要求必须检测到 Heading 样式。

成功时：

- `status = completed`；
- `docx_status = valid`；
- `result_source = word_keep_source_format`；
- `chat_url` 保存当前对话链接。

## 10. 任务状态和手动修改

| 状态 | 含义 |
|---|---|
| `pending` | 等待处理 |
| `processing` | 已领取，网页正在处理 |
| `completed` | Word 已生成并通过本地检查 |
| `needs_review` | 自动化失败、结果需检查或需要重新执行 |
| `failed` | 本地输入缺失等原因导致无法继续 |

可勾选一个或多个任务，通过下拉框批量设为：

- `pending`；
- `needs_review`；
- `completed`；
- `failed`。

默认选择是 `completed`。

手动设为 `completed` 时，系统会先检查对应 Word 是否存在、可打开且大小合格。如果任务目录只有一个非临时 `.docx`，系统会先把它改名为论文名再检查。Word 的 `~$` 临时锁文件会忽略。

修改真正正在运行的 `processing` 任务可能与网页流程冲突，界面会要求二次确认。若任务只是因为浏览器中断而卡在 `processing`，可在确认没有自动化仍在运行后手动改为 `pending`。

## 11. 删除论文和任务

1. 勾选一个或多个任务。
2. 点击“删除所选”。
3. 阅读确认提示。
4. 确认后，系统会删除：

   - PDF；
   - Word；
   - 整个任务输出目录；
   - `progress.csv` 中的任务记录。

删除是永久操作，不进入回收站。建议先备份重要结果。

如果某个文件正在 Word 或其他程序中打开，删除可能失败；失败任务会保留并显示编号。关闭占用文件后重试。

如果你已经在资源管理器中手动删除了整个任务输出目录，刷新列表时本地服务会自动移除相应的残留任务记录。

## 12. `progress.csv` 字段说明

`progress.csv` 使用 UTF-8-SIG 编码，可直接用 Excel 打开。

| 字段 | 含义 |
|---|---|
| `task_id` | 三位任务编号，如 `001` |
| `paper_name` | 原始论文文件名 |
| `source_path` | `output` 内 PDF 的相对路径 |
| `output_folder` | 当前任务输出目录 |
| `status` | 任务状态 |
| `chat_url` | ChatGPT 对话 URL |
| `docx_status` | `missing`、`valid` 或 `invalid` |
| `result_source` | Word 生成来源 |
| `quality_check` | 本地检查结果 |
| `error_message` | 最近失败原因 |
| `updated_at` | 最近更新时间 |

不要在扩展运行期间同时用 Excel 保存 `progress.csv`，否则 Excel 的文件锁可能阻止状态更新。

## 13. 每天使用时的最短启动流程

每次重新开机后：

1. 打开 PowerShell；
2. 进入项目目录；
3. 运行：

   ```powershell
   py workflow/local_server.py
   ```

4. 打开 Chrome；
5. 登录 ChatGPT；
6. 点击扩展图标；
7. 拖入 PDF；
8. 选择任务；
9. 点击“保存并开始”。

`.workflow-token` 通常只需第一次粘贴，扩展会保存在 Chrome 本地存储中。

## 14. 停止任务

点击扩展中的“停止”：

- 阻止继续领取下一篇；
- 将尚未真正开始的当前领取任务释放回 `pending`；
- 不会强制中止 ChatGPT 已经开始生成的回答；
- 不会关闭 Word 或浏览器。

如果当前回答仍在网页生成，请等待它结束或手动停止 ChatGPT 生成。

## 15. 可选命令行工具

正常使用只需要扩展和 `local_server.py`。以下工具主要用于诊断。

检查服务：

```powershell
Invoke-RestMethod http://127.0.0.1:8765/api/health
```

检查已有 Word：

```powershell
py workflow/check_results.py
```

检查单个任务，并且不要求内置 Heading 样式：

```powershell
py workflow/check_results.py --task-id 001 --min-headings 0
```

命令行领取下一篇 pending 任务：

```powershell
py workflow/get_next_task.py
```

重新领取 `needs_review`：

```powershell
py workflow/get_next_task.py --include-needs-review
```

这些命令不会替代 Chrome 网页自动化；尤其不要在扩展正在运行时再用 `get_next_task.py` 领取任务，否则可能产生状态冲突。

`finish_task.py` 和 `finish_text.py` 是旧的手工归档辅助脚本，不属于当前推荐的“复制回答 → Word 原格式粘贴”自动流程。

## 16. 常见问题与排查

### 16.1 扩展显示无法连接本地服务

检查：

```powershell
Invoke-RestMethod http://127.0.0.1:8765/api/health
```

如果失败：

- 确认 `local_server.py` 终端仍开着；
- 确认端口是 `8765`；
- 确认扩展地址是 `http://127.0.0.1:8765`；
- 检查防火墙或安全软件是否拦截本机回环连接；
- 重新启动本地服务。

### 16.2 提示令牌无效

- 打开项目根目录 `.workflow-token`；
- 复制整行，不要带多余空格；
- 粘贴到扩展令牌框；
- 如果更换了项目副本，使用新项目目录中的令牌。

不要把 `.workflow-token` 发给其他人。

### 16.3 点击扩展图标没有窗口

- 打开 `chrome://extensions/`；
- 确认扩展已启用；
- 点击“重新加载”；
- 确认所需权限已允许；
- 查看扩展页面中的“错误”；
- 再次点击图标。

### 16.4 扩展版本没有更新

必须在 `chrome://extensions/` 点击“重新加载”，然后关闭旧任务窗口并重新打开。只刷新 ChatGPT 页面不会更新后台 service worker 和任务窗口脚本。

### 16.5 PDF 上传失败

检查：

- 文件扩展名是 `.pdf`；
- 文件不是 0 字节；
- 单个文件不超过 200 MB；
- 文件确实以 PDF 格式保存；
- 是否已有同名任务；
- 本地服务是否运行；
- `output` 是否可写。

### 16.6 任务一直是 processing

可能原因：

- ChatGPT 仍在生成；
- 页面结构变化；
- 浏览器或扩展被重载；
- Word 自动化仍在进行；
- `progress.csv` 被其他程序锁定。

先查看 `error_message`。确认没有自动化仍在运行后，可以停止队列并把任务手动改为 `pending`。

### 16.7 没有生成 Word

常见原因：

- 没有安装 Microsoft Word 桌面版；
- ChatGPT 的复制按钮没有正确更新剪贴板；
- Word 粘贴内容与当前回答不一致；
- Word 文件被打开并锁定；
- 回答太短；
- Word 文件小于 20 KB；
- Chrome 没有允许 `debugger` 或 `clipboardRead` 权限。

系统会拒绝归档不匹配的 Word，而不是把错误内容当成成功结果。把任务设为 `needs_review` 或 `pending` 后可以重新执行。

### 16.8 Word 只包含旧剪贴板内容

当前版本在复制前会自动激活 ChatGPT 标签页，在启动 Word 前校验剪贴板，Word 保存后再校验正文。如果仍发生，确认扩展版本是 `1.1.1` 并重新加载扩展。不要在任务完成瞬间手动复制其他内容。

### 16.9 只复制了表格或代码

扩展优先选择 ChatGPT 的 `copy-turn-action-button`，并排除回答内部的“复制表格”和“复制代码”。如果 ChatGPT 更新了页面结构，可能需要更新 `content.js` 的选择器。

### 16.10 新开的 ChatGPT 标签页也开始执行任务

扩展只控制启动时锁定的标签页。如果出现异常：

- 点击“停止”；
- 关闭多余 ChatGPT 标签页；
- 重新加载扩展；
- 打开一个空白 ChatGPT 标签页后重新开始。

### 16.11 Chrome 显示正在被调试

这是可信点击复制按钮所需的正常提示。扩展只在发送复制按钮点击时短暂连接 Chrome 调试协议，随后立即断开。

### 16.12 删除任务失败

- 关闭对应 Word；
- 关闭正在预览该 PDF 的程序；
- 停止正在处理的队列；
- 确认输出目录没有被其他程序锁定；
- 再次点击“删除所选”。

## 17. 安全与隐私

- 本地服务只监听 `127.0.0.1`，不对局域网或公网开放；
- 扩展请求必须携带 `.workflow-token`；
- PDF 和 Word 保存在本机 `output`；
- PDF 会上传到你登录的 ChatGPT 网页会话，因此仍受该账号和组织的数据策略约束；
- 本项目不保存 ChatGPT 密码；
- 本项目不需要 OpenAI API Key；
- 不要把 `.workflow-token`、论文或输出目录提交到公开仓库；
- 如果论文含敏感或受限数据，请先确认组织政策允许上传到当前 ChatGPT 工作区。

## 18. 已知限制

- 只能串行处理，不能并行运行多个受控 ChatGPT 标签页；
- ChatGPT 网页结构变化后，按钮和输入框选择器可能需要更新；
- 验证码、登录确认、使用额度限制需要人工处理；
- 必须安装 Windows 桌面版 Word；
- 自动检查只能确认文件结构和回答一致性，不能保证论文分析内容在学术上完全正确；
- Word 的网页原格式粘贴不一定把视觉标题映射为 Word 内置 Heading 样式；
- 删除所选是永久删除；
- 运行期间不应手动操作受控 ChatGPT 标签页。

## 19. 推荐的备份内容

至少备份：

- `output/`；
- `progress.csv`；
- `prompt/analysis_prompt.md`。

通常不需要备份 `.workflow-token`，丢失后可停止服务、删除该文件并重新启动，以生成新令牌；随后在扩展中更新令牌。

## 20. 快速检查清单

开始批处理前逐项确认：

- [ ] Windows 和 Chrome 可正常使用；
- [ ] ChatGPT 已登录且能上传 PDF；
- [ ] Microsoft Word 桌面版已安装；
- [ ] Python 和依赖已安装；
- [ ] `analysis_prompt.md` 已填写；
- [ ] `local_server.py` 正在运行；
- [ ] 扩展版本为 `1.1.1`；
- [ ] 扩展令牌正确；
- [ ] 已打开一个空白 ChatGPT 标签页；
- [ ] 没有打开将被替换的任务 Word；
- [ ] 已选择正确任务；
- [ ] 处理期间不会手动操作受控标签页。
