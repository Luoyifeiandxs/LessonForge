# LessonForge —— 讲课视频生成 Agent 设计文档

- 日期：2026-09-20
- 状态：已与需求方逐节评审通过，待实施
- 形态：本地自用工具（Windows），单用户，无部署

## 1. 背景与目标

参考 B 站视频《神经科学：探索大脑》丨第一章（BV1yxkqBEEG8，"学科学巨著"系列）：约 56 分钟的讲课视频，画面为全屏幻灯片 + AI 配音 + 逐句字幕，无真人出镜。

目标：做一条本地流水线，把"教材/书籍章节"加工成这类视频，且在生成与成片之间保留**人工校对**环节。

### 流程（用户定义）

```
输入（教材 PDF/电子文本；可选：现成 PPT、现成教案）
  → ① 生成 PPT + 每页讲稿
  → ② 人工校对（看图、改稿、单页重生成、试听）
  → ③ 自动渲染：逐页幻灯片 + TTS 配音 + 字幕 → MP4
```

### 明确不做（YAGNI）

- 账号系统、多用户、消息队列
- pptx 动画还原、自动翻译
- BGM / 片头片尾 / 数字人（只预留数据与渲染槽位，不实现，见 §10）

## 2. 核心设计决策（评审结论记录）

| # | 决策点 | 结论 | 理由 |
|---|--------|------|------|
| 1 | 成片形态 | 纯"PPT+配音+字幕"，渲染器按分层时间线架构预留升级位 | B（片头/BGM）/C（数字人）升级时零重构 |
| 2 | PPT 路线 | 混合：AI 新生成的用 Marp markdown；上传的 pptx 走 LibreOffice 转换渲染 | 各取所长；转换失败有重排兜底 |
| 3 | 定位 | 本地自用，界面能用就行，单渲染并发 | 砍掉账号/队列/多租户复杂度 |
| 4 | 内容来源 | 有源材料（教材章节喂给 LLM 防幻觉）；文字版与扫描版 PDF 都存在 | 摄取层可插拔，MinerU 纳入默认路线 |
| 5 | 教案角色 | 可选参考输入，不是必经中间产物 | 流程更短 |
| 6 | TTS | edge-tts 起步（免费、字级时间戳），外包 TTSProvider 适配器 | 以后可换 Azure/火山/本地模型 |
| 7 | 编排框架 | LangChain（LCEL）负责所有 LLM 调用；流程编排用普通 Python 状态机；不用 LangGraph/CrewAI | 流程是线性五段+一个人工中断，不确定性全在 chain 内部；理由与升级触发器见 §11 |
| 8 | 大纲确认关卡 | 开启：先生成大纲→人工确认→再生成全部 | 页数骨架先定对，避免大面积返工 |
| 9 | 校对粒度 | 幻灯片/讲稿按页独立存、独立重生成、TTS 按页缓存 | 改 1 页只重跑 1 页 |

## 3. 总体架构

```
┌──────────────── 本地 Web 界面（FastAPI + Jinja2 + HTMX，无前端构建链）────────────────┐
│ 项目列表页 │ 项目工作台页（材料区/生成区/校对区/渲染区）                              │
└────────────────────────────────────┬───────────────────────────────────────────────┘
                                     │
┌────────────────────────────────────▼───────────────────────────────────────────────┐
│ FastAPI 后端 —— 线性五阶段状态机，每阶段产物落盘、幂等可续跑                         │
│                                                                                     │
│ ① 摄取   MaterialParser：PDF→章节markdown+插图；LangChain Splitter/Embedding→Chroma │
│ ② 生成   LangChain LCEL 三链：大纲 → 逐页幻灯片(Marp) → 逐页讲稿                     │
│ ③ 校对   人工一页一屏过（interrupt 点，无状态标记，见 §7）                           │
│ ④ 配音   TTSProvider(edge-tts)：逐页 mp3 + 字级时间戳，按 hash 缓存                  │
│ ⑤ 渲染   marp 截图 / LibreOffice 转换 → ffmpeg 分段→拼接→烧字幕 → MP4               │
└─────────────────────────────────────────────────────────────────────────────────────┘
```

### 技术栈

| 环节 | 组件 | 说明 |
|------|------|------|
| LLM | LangChain + langchain-openai | OpenAI 兼容 base_url，可配 DeepSeek/Kimi/OpenAI；结构化输出用 Pydantic + `with_structured_output` |
| PDF 摄取 | PyMuPDFParser（默认轻量）/ MinerUParser（质量优先） | 见 §5.1 |
| 切分/检索 | LangChain RecursiveCharacterTextSplitter + Chroma（本地持久化） | embedding 为独立 OpenAI 兼容端点（DeepSeek 不提供 embedding；默认 OpenAI 或本地服务，`LF_EMB_*` 环境变量，未配则跳过检索增强） |
| 幻灯片 | Marp（markdown→PNG，自定义 CSS 主题） | `marp slides.md --images png --allow-local-files`；LLM 写 markdown 比裸 HTML 可靠 |
| pptx 转换 | LibreOffice headless → PDF → PyMuPDF 渲染 PNG；python-pptx 抽文本 | 全现成组件 |
| TTS | edge-tts（WordBoundary 字级事件） | 外套 TTSProvider 抽象 |
| 合成 | ffmpeg（分段 `-loop 1` → concat 流复制 → 一遍编码烧字幕） | ASS 烧录 + 附带 .srt |
| 界面 | FastAPI + Jinja2 + HTMX | 服务端渲染+局部刷新，1 秒轮询长任务进度 |
| 状态 | project.json + 文件夹产物，无数据库 | 单用户自用；阶段幂等天然支持续跑 |

## 4. 数据模型

### 4.1 项目文件夹

```
projects/<项目id>/
├── project.json            # 状态机 + 配置 + 每页元数据（唯一状态源）
├── material/
│   ├── raw/                # 上传的 PDF / pptx / 教案原文
│   ├── parsed/             # 章节markdown + images/（原书插图，MinerU 产出带图注）
│   └── index/              # Chroma 向量索引
├── deck/
│   ├── outline.json        # 大纲（人工确认后的版本）
│   ├── slides.md           # Marp 全稿（幻灯片主产物）
│   ├── parts/              # slide-01.md ... 逐页 Marp 片段（单页重生成的单元）
│   ├── theme.css           # 自制主题（2~3 个版式：标题页/要点页/图文页）
│   └── png/                # slide-01.png ...
├── scripts/                # scripts/slide-01.md 逐页讲稿（校对以这里为准）
├── audio/
│   ├── slide-01.mp3        # 逐页配音（稳定文件名，重配音直接覆盖）
│   ├── slide-01.json       # sidecar：{duration, word_boundaries[]}，字幕对齐唯一来源
│   └── cache/              # hash(讲稿+音色) → mp3/json 的 TTS 缓存
├── timeline.json           # 渲染唯一输入：每页图片/音频/时长/字幕cue + 预留槽位
└── output/
    ├── final.mp4
    ├── subtitles.ass
    └── subtitles.srt
```

### 4.2 project.json 关键结构

```jsonc
{
  "id": "neuroscience-ch1",
  "title": "《神经科学：探索大脑》第一章",
  "config": {
    "parser": "mineru-local | mineru-api | pymupdf",
    "llm": {"base_url": "...", "model": "..."},   // OpenAI 兼容
    "voice": "zh-CN-YunxiNeural",
    "resolution": [1920, 1080],
    "page_gap_sec": 0.4,
    "outline_gate": true
  },
  "stages": {                    // 阶段级状态：pending|running|done|failed
    "ingest": "done", "outline": "done", "slides": "done",
    "audio": "done", "render": "pending"
  },
  "slides": [                    // 页级元数据，失败/重试定位到页
    {
      "index": 5,
      "title": "神经元的解剖结构",
      "png": "deck/png/slide-05.png",
      "script": "scripts/slide-05.md",
      "audio_hash": "ab12…",   // 与讲稿+音色 hash 对应；不等则需重配
      "audio_duration": 42.3,
      "error": null
    }
  ]
}
```

### 4.3 timeline.json

```jsonc
{
  "resolution": [1920, 1080],
  "segments": [
    {
      "index": 5, "png": "deck/png/slide-05.png",
      "audio": "audio/slide-05.mp3", "duration": 42.7,   // 音频时长 + page_gap
      "subtitle_cues": [{"start": 0.0, "end": 2.31, "text": "神经元是大脑的基本单位。"}],
      "overlay": null,   // 预留槽位（数字人/角标），本期恒为 null
      "bgm": null        // 预留槽位，本期恒为 null
    }
  ]
}
```

## 5. 流水线阶段设计

### 5.1 摄取（MaterialParser 可插拔）

```
MaterialParser（接口）→ {章节列表: [{标题, markdown 文本, 图片清单}]}
├─ PyMuPDFParser   默认轻量路径：文字版 PDF；章节切分优先取 PDF 书签(TOC)，无书签则 LLM 切分兜底
└─ MinerUParser    质量路径：mineru CLI（本地 GPU）或 mineru.net API
                   产出 markdown（标题层级/公式 LaTeX/表格）+ images/（插图+图注）
                   配置项 parser 三选一：pymupdf | mineru-local | mineru-api
```

摄取完成后：RecursiveCharacterTextSplitter 切块 → embedding → Chroma 入库（`material/index/`）。扫描版 PDF 必须走 MinerU（PyMuPDF 提不出文字）。

### 5.2 生成（LangChain LCEL 三链，全部结构化输出）

1. **大纲链**：材料章节（+可选教案原文）→ `{slides: [{title, points, figure_hint, teaching_goal}]}` → **人工确认关卡**（工作台内可编辑大纲文本，确认后写盘 `outline.json`；确认前不生成任何幻灯片/讲稿）。页数在此定死。
2. **逐页幻灯片链**：按大纲逐页生成 Marp markdown；上下文 = 大纲该项 + 前 2 页标题 + 该页检索到的源材料块（含 `figure_hint` 指向的插图路径）→ 追加拼入 `slides.md`（页间 `---` 分隔）→ marp 出 PNG。
3. **逐页讲稿链**：逐页生成口语化讲稿 150~300 字（对应 40~70s 音频），同样带检索上下文 → 写 `scripts/slide-0N.md`。约束：数字一律写中文汉字（如「一百二十」），避免 TTS 误读——该约束写入幻灯链与讲稿链两者的 system prompt。

三链均为"读产物→写产物"的纯函数：不碰 HTTP、不碰全局状态（LangGraph 升级的前提纪律，见 §11）。

### 5.3 配音（TTSProvider 适配器）

```python
class TTSProvider(Protocol):
    async def synth(self, text: str, voice: str, out_path: str) -> AudioResult  # mp3 + 字级边界
# 本期唯一实现 EdgeTTSProvider；预留 AzureTTSProvider / LocalGPTSoVITSProvider 位置
```

- **句切分**：中文按 `。！？；` 切句（纯函数，单测覆盖）
- **缓存**：`hash(讲稿文本 + voice)` 命中 `audio/cache/` 则跳过合成；讲稿未变的页永远不重配
- 讲稿保存事件 → 该页 hash 变化 → 后台自动重配该页（见 §7）

### 5.4 渲染（ffmpeg 两遍法）

1. 逐页分段：`ffmpeg -loop 1 -i slide-0N.png -i slide-0N.mp3 -t <duration>` → 段 mp4（h264/yuv420p，静态图编码快）
2. concat demuxer 流复制拼接（秒级，无重编码）
3. 一遍编码烧字幕：全局 `subtitles.ass`（页内句级 cue 由字级边界聚合，全局时间 = 前序页时长累加 + page_gap）
4. 附带导出 `.srt`；BGM 轨与 overlay 槽位本期不渲染，仅存在于 timeline.json

### 5.5 上传 pptx 路径

- 常规：`soffice --headless --convert-to pdf` → PyMuPDF 渲 PNG → python-pptx 抽每页文字作为讲稿生成的上下文
- 兜底：转换失败 → 自动切"重排模式"：抽出的文字走 5.2 生成链，按自家 Marp 主题重新排版（丢动画保内容）

## 6. 字幕时间轴机制

```
讲稿 → 句切分 → edge-tts WordBoundary（字级起止）
     → 句级 cue = 该句首字 start ~ 末字 end（页内相对）
     → 全局 = Σ 前序页音频时长 + page_gap
     → timeline.json + subtitles.ass/.srt
```

逐句上屏（一句一屏），不做整页糊屏；不引入强制对齐工具。默认值：1080p、烧录+软字幕双出、页间隔 0.4s、默认音色 zh-CN-YunxiNeural，均可在 project.json 配置。

一轮校对的典型成本 = 改动页数 × 单页 TTS（秒级）+ 秒级拼接 + 一遍字幕编码，56 分钟成片约几分钟。

## 7. Web 界面

两个页面，HTMX 局部刷新 + 1s 轮询长任务进度。

### 7.1 项目列表页

新建项目（标题、上传材料、选 parser/音色/模型）、项目卡片（标题、阶段状态徽章、进入按钮）。

### 7.2 项目工作台页（四个区块，单页滚动）

1. **材料区**：章节树、插图数量、[重新解析]
2. **生成区**：[生成大纲] → 大纲以逐页文本块编辑（每块 = 标题+要点，支持增删/改写/排序，非裸 JSON）→ [确认大纲并生成]，逐页进度条
3. **校对区（一页一屏线性模式）**：

```
◀ 上一页   第 5 / 80 页   下一页 ▶        [跳到第 __ 页]   ←→ 键盘翻页
┌──────────────────────────────────────┐
│        幻灯片大图（占主要面积）        │
├──────────────────────────────────────┤
│ 讲稿 textarea（可直接编辑）            │
│ [保存]（后台自动重配该页配音） [▶ 试听] │
│ [重新生成讲稿]  [重新生成本页幻灯片]   │
├──────────────────────────────────────┤
│ …翻页到底 → [生成视频]                │
└──────────────────────────────────────┘
```

设计原则：**无逐页审批动作、无状态徽章**。翻页即校对；讲稿保存 → hash 变化 → 该页后台重配音，音频永远最新，无需"过期"提醒。

4. **渲染区**：[生成视频] → 进度 → 成片内嵌播放 + 下载 MP4/SRT

### 7.3 后台任务模型

- 进程内 asyncio 任务队列，**同一时间只跑一个重活**；批量生成/批量 TTS/渲染、以及校对触发的单页配音，全部经同一队列串行执行
- 状态写回 project.json；应用重启时未完成的 running 任务标记为 failed(可重试)，UI 单击续跑（产物幂等，重跑只补缺）

## 8. 错误处理

1. **失败定位到页**：LLM/TTS/截图哪页挂了标哪页，UI 单页重试
2. **LLM 结构化输出失败**：output parser 带错误信息自动重试 2 次；仍失败 → 原始文本进校对界面人工救，流水线不停摆
3. **依赖体检**：`python -m app doctor` 检查 ffmpeg / LibreOffice / marp(npx，需 Node) / GPU，缺失给出安装指引，不让错误拖到渲染时才爆
4. **pptx 转换失败**：自动切重排模式（§5.5）
5. **edge-tts 网络失败**：指数退避重试；缓存保证未变文本不重复请求

## 9. 测试策略

- **单元**：句切分、时间轴聚合（字级→句级→全局）、project.json 状态流转、TTS 缓存 hash、幻灯片 JSON→Marp markdown 拼装（golden 断言）
- **集成**：3 页 golden mini 项目（mock LLM + 罐头音频）跑通 ①→⑤ 全链，逐步断言产物存在与内容正确
- **手动 smoke 清单**：真实 LLM + 真实 TTS 跑一个真实 PDF 章节到成片，不进自动化

## 10. 升级路径（本期不实现，仅保证不重构）

| 升级 | 改动 | 依赖的预留 |
|------|------|-----------|
| B：片头/片尾/BGM/章节角标 | 小：时间线首尾插段；启用 bgm 槽（ffmpeg amix + 配音 ducking）；角标走 overlay | timeline.json 槽位、两遍渲染结构 |
| C1：角落静态/循环头像（无口型） | 小：overlay 槽贴循环素材 | overlay 槽位 |
| C2：口型同步数字人 | 较大：新增"讲稿+音频→人物视频段"生成模块，产物进 overlay | 分段时间线结构不变 |

## 11. 编排框架决策：LangChain LCEL + Python 状态机

**不用 CrewAI**：流程是固定流水线，不是多角色协作；角色扮演徒增成本、随机性与调试难度。

**暂不用 LangGraph**：流程为线性五段 + 恰一个人工中断（校对），LangGraph 的 checkpoint/interrupt 与 project.json + asyncio 队列能力重叠；引入会造成磁盘产物与图状态两份真相。

**升级触发器**（满足任一再引入 LangGraph）：

1. 新增自主环节（如自动联网补材料）
2. 出现跨阶段自我修正循环（如音频超时自动改写讲稿压时长）
3. 人工中断点 ≥ 2 个

由于 5.2 三链均为纯函数，届时包成 graph node 是机械改造（约半天），无沉没成本。

## 12. 环境依赖

| 依赖 | 用途 | 安装 |
|------|------|------|
| Python 3.11+ | 主运行时 | 已有 |
| Node.js（npx marp） | Marp CLI 截图 | 已有 |
| ffmpeg | 合成渲染 | winget/scoop，doctor 检查 |
| LibreOffice | pptx→PDF | 官网安装，doctor 检查路径 |
| MinerU（可选） | 质量摄取 | `pip install mineru`（本地 GPU）或走 mineru.net API，免装 |

LLM：任一 OpenAI 兼容端点（默认 DeepSeek，配置可换 Kimi/OpenAI）；Embedding：独立 OpenAI 兼容端点（默认 OpenAI，可接本地服务；未配置则跳过检索增强）。
