# LessonForge 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 本地自用的"教材 → PPT+逐页讲稿 → 人工校对 → MP4 讲课视频"流水线工具。

**Architecture:** FastAPI 单体 + Jinja2/HTMX 界面；五个线性阶段（摄取/生成/校对/配音/渲染）由普通 Python 状态机编排，LangChain LCEL 只负责 LLM 调用；一切产物按页落盘，project.json 是唯一状态源。

**Tech Stack:** Python 3.11+，FastAPI，LangChain(langchain-openai)，Chroma，PyMuPDF，MinerU(可选)，python-pptx，Marp CLI(npx)，edge-tts，ffmpeg，Jinja2+HTMX。

**Spec:** `docs/specs/2026-09-20-lecture-video-agent-design.md`（本计划的唯一需求来源）

## Global Constraints（摘自 spec，逐字生效）

- Python ≥ 3.11；主平台 Windows；不用数据库，状态 = project.json
- 分辨率默认 1920×1080；页间隔默认 0.4s；默认音色 `zh-CN-YunxiNeural`（均可配置）
- 字幕逐句上屏；烧录进画面 + 附带 .srt
- TTS 仅 edge-tts，但必须包在 `TTSProvider` 适配器后；缓存 key = `sha1(voice::讲稿文本)[:16]`
- 编排只用 LangChain LCEL + Python 状态机；禁止引入 LangGraph / CrewAI / 自主 Agent 循环
- 每个阶段函数是"读产物→写产物"的纯函数：不碰 HTTP、不碰全局状态
- 失败定位到"页"；LLM 结构化输出失败自动带错误重试 2 次，仍失败走原始文本兜底，流水线不停摆
- 同一时间只跑一个重活；所有任务（含单页配音）走同一 asyncio 串行队列
- 界面 = 2 个页面（项目列表 / 项目工作台），HTMX 局部刷新；校对区为一页一屏线性模式，无逐页审批动作、无状态徽章
- 大纲确认关卡默认开启；幻灯片/讲稿按页独立存储、独立重生成
- 讲稿生成提示词必须要求"数字用中文汉字书写"（TTS 朗读准确性）
- 喂给大纲链的每个 chapter 文本截断到 20000 字符以内（上下文预算）

## 文件结构总览

```
pyproject.toml                 依赖与 pytest 配置
app/
  config.py                    全局路径(ROOT/PROJECTS_DIR) + 环境变量默认配置
  models.py                    全部 Pydantic 数据模型（spec §4）
  store.py                     项目文件夹/原子读写/状态存取
  jobs.py                      asyncio 串行任务队列 + 重启恢复
  doctor.py                    依赖体检 CLI（python -m app.doctor）
  main.py                      FastAPI 应用工厂、静态资源、路由挂载
  pipeline/
    ingest.py                  MaterialParser 接口 + PyMuPDFParser + MinerUParser
    index.py                   切分 + Chroma 索引 + 检索
    llm.py                     ChatOpenAI / OpenAIEmbeddings 工厂
    chains.py                  大纲/幻灯片/讲稿三链 + ask_structured 重试
    deck.py                    theme.css、slides.md 组装、marp 截图
    pptx.py                    pptx→PNG(LibreOffice)、文本抽取、重排兜底
    tts.py                     TTSProvider 协议 + EdgeTTSProvider + 缓存
    timeline.py                句切分、字级→句级对齐、timeline.json、ASS/SRT
    render.py                  ffmpeg 分段→拼接→烧字幕
    orchestrator.py            五阶段编排函数（run_ingest 等）
  web/
    routes_projects.py         项目列表页 + 新建
    routes_workbench.py        工作台四区全部端点
    templates/                 base/projects/workbench + partials
    static/app.css
tests/
  conftest.py                  tmp 项目根 fixture
  fakes.py                     FakeStructuredLLM / FakeTTSProvider / FakeEmbeddings
  test_*.py                    每任务配套测试
docs/smoke.md                  手动全链 smoke 清单
```

任务顺序即依赖顺序：1→2→3 为地基；4~12 相互独立（建议按序）；13→14 编排；15 体检；16→17 界面；18 收尾。

---

### Task 1: 项目脚手架与全局配置

**Files:**
- Create: `pyproject.toml`, `app/__init__.py`, `app/config.py`, `tests/conftest.py`, `.gitignore`

**Interfaces:**
- Produces: `app.config.ROOT: Path`、`app.config.PROJECTS_DIR: Path`、`app.config.project_dir(pid) -> Path`、`app.config.default_project_config() -> dict`（供新建项目用，值来自环境变量）

- [ ] **Step 1: 写 pyproject.toml**

```toml
[project]
name = "lessonforge"
version = "0.1.0"
requires-python = ">=3.11"
dependencies = [
  "fastapi>=0.110", "uvicorn[standard]>=0.29", "jinja2>=3.1", "python-multipart>=0.0.9",
  "pydantic>=2.6", "langchain>=0.2", "langchain-openai>=0.1", "langchain-chroma>=0.1",
  "chromadb>=0.5", "langchain-text-splitters>=0.2", "pymupdf>=1.24", "python-pptx>=0.6",
  "edge-tts>=6.1", "httpx>=0.27",
]
[project.optional-dependencies]
dev = ["pytest>=8", "pytest-asyncio>=0.23", "pillow>=10"]
[tool.pytest.ini_options]
asyncio_mode = "auto"
testpaths = ["tests"]
```

- [ ] **Step 2: 安装依赖**

Run: `pip install -e ".[dev]"`
Expected: 安装成功无报错

- [ ] **Step 3: 写 app/config.py 与 .gitignore**

```python
# app/config.py
import os
from pathlib import Path

ROOT = Path(os.environ.get("LESSONFORGE_ROOT", Path.home() / "LessonForge"))
PROJECTS_DIR = ROOT / "projects"

def project_dir(pid: str) -> Path:
    return PROJECTS_DIR / pid

def default_project_config() -> dict:
    """环境变量兜底的 LLM/Embedding 配置（OpenAI 兼容端点，可任意换供应商）"""
    return {
        "parser": os.getenv("LF_PARSER", "pymupdf"),
        "llm": {
            "base_url": os.getenv("LF_LLM_BASE_URL", "https://api.deepseek.com/v1"),
            "api_key": os.getenv("LF_LLM_API_KEY", ""),
            "model": os.getenv("LF_LLM_MODEL", "deepseek-chat"),
        },
        "embedding": {
            "base_url": os.getenv("LF_EMB_BASE_URL", "https://api.openai.com/v1"),
            "api_key": os.getenv("LF_EMB_API_KEY", ""),
            "model": os.getenv("LF_EMB_MODEL", "text-embedding-3-small"),
        },
        "voice": os.getenv("LF_VOICE", "zh-CN-YunxiNeural"),
        "resolution": [1920, 1080],
        "page_gap_sec": 0.4,
        "outline_gate": True,
    }
```

`.gitignore`: `__pycache__/`, `*.egg-info/`, `.pytest_cache/`, `projects/`, `.venv/`

- [ ] **Step 4: 写 tests/conftest.py（tmp 项目根）**

```python
# tests/conftest.py
import pytest
from app import config

@pytest.fixture(autouse=True)
def tmp_root(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "ROOT", tmp_path)
    monkeypatch.setattr(config, "PROJECTS_DIR", tmp_path / "projects")
    return tmp_path
```

注意：后续所有 store/pipeline 模块读取路径时**必须 `from app import config` 后取 `config.PROJECTS_DIR`**（不能 `from app.config import PROJECTS_DIR`），否则 monkeypatch 失效。此为全项目硬性约定。

- [ ] **Step 5: 冒烟测试 tests/test_scaffold.py**

```python
def test_paths_are_under_root(tmp_root):
    from app import config
    assert config.project_dir("abc").is_relative_to(tmp_root)
```

Run: `pytest tests/test_scaffold.py -v` → Expected: PASS
- [ ] **Step 6: Commit** `git add -A && git commit -m "chore: 项目脚手架与全局配置"`

---

### Task 2: 数据模型（spec §4 全量）

**Files:**
- Create: `app/models.py`, `tests/test_models.py`

**Interfaces:**
- Produces（后续所有任务引用，字段名逐字一致）:
  - `ParserKind(str,Enum)`: `pymupdf|mineru_local|mineru_api`
  - `StageStatus(str,Enum)`: `pending|running|done|failed`
  - `LLMConfig(base_url:str="", api_key:str="", model:str="")`；`EmbeddingConfig(base_url:str="", api_key:str="", model:str="")`（空配置合法：建项目时可暂不填，doctor 会报告未配置）
  - `ProjectConfig(parser:ParserKind=ParserKind.pymupdf, llm:LLMConfig=LLMConfig(), embedding:EmbeddingConfig=EmbeddingConfig(), voice:str="zh-CN-YunxiNeural", resolution:list[int]=[1920,1080], page_gap_sec:float=0.4, outline_gate:bool=True)`
  - `Chapter(title:str, markdown:str, images:list[str]=[])`
  - `OutlineSlide(title:str, points:list[str], figure_hint:str|None=None, teaching_goal:str="")`；`Outline(slides:list[OutlineSlide])`
  - `SlidePart(marp_md:str)`；`ScriptPart(narration:str)`
  - `WordBoundary(offset:float, duration:float, text:str)`（秒）
  - `AudioResult(path:str, duration:float, word_boundaries:list[WordBoundary]=[])`
  - `SubtitleCue(start:float, end:float, text:str)`
  - `Segment(index:int, png:str, audio:str, duration:float, subtitle_cues:list[SubtitleCue], overlay:None=None, bgm:None=None)`；`Timeline(resolution:list[int], segments:list[Segment])`
  - `SlideMeta(index:int, title:str="", png:str|None=None, script:str|None=None, audio_hash:str|None=None, audio_duration:float|None=None, error:str|None=None)`
  - `ProjectState(id:str, title:str, created_at:str, config:ProjectConfig, stages:dict[str,str], slides:list[SlideMeta], errors:dict[str,str]={})`；`new_stages() -> dict`（五键 `ingest/outline/slides/audio/render` 全 `pending`）

- [ ] **Step 1: 写实现（模型为声明式，直接写全）**——按上面 Interfaces 逐字实现，`new_stages()` 返回 `{"ingest":"pending","outline":"pending","slides":"pending","audio":"pending","render":"pending"}`

- [ ] **Step 2: 写测试 tests/test_models.py**

```python
import app.models as M

def test_new_stages_keys():
    assert list(M.new_stages()) == ["ingest", "outline", "slides", "audio", "render"]

def test_defaults_match_spec():
    c = M.ProjectConfig(parser=M.ParserKind.pymupdf,
                        llm=M.LLMConfig(base_url="u", api_key="k", model="m"),
                        embedding=M.EmbeddingConfig(base_url="u", api_key="k", model="e"))
    assert c.voice == "zh-CN-YunxiNeural" and c.resolution == [1920, 1080] and c.page_gap_sec == 0.4
    assert c.outline_gate is True

def test_segment_reserved_slots_default_none():
    seg = M.Segment(index=1, png="a.png", audio="a.mp3", duration=1.0, subtitle_cues=[])
    assert seg.overlay is None and seg.bgm is None
```

- [ ] **Step 3: Run** `pytest tests/test_models.py -v` → PASS
- [ ] **Step 4: Commit** `git add -A && git commit -m "feat: 数据模型（spec §4）"`

---

### Task 3: 项目仓库 store.py（原子读写）

**Files:**
- Create: `app/store.py`, `tests/test_store.py`

**Interfaces:**
- Produces:
  - `create_project(title: str, config: dict) -> ProjectState`（建目录树 + 写 project.json）
  - `load_project(pid: str) -> ProjectState`；`save_project(state: ProjectState) -> None`（原子：tmp + `os.replace`）
  - `list_projects() -> list[ProjectState]`（按 created_at 倒序）
  - `read_text(path: Path) -> str`、`write_text(path: Path, s: str) -> str`（write 返回绝对路径字符串，自动建父目录，同样原子）
  - `sl(n: int) -> str`：页文件名工具 `"slide-%02d"` 前缀生成器，`sl(5)=="slide-05"`；`f_slides(pid)`/`f_script(pid,n)`/`f_audio(pid,n)`/`f_audio_meta(pid,n)`/`f_png(pid,n)` 等路径函数

- [ ] **Step 1: 写测试 tests/test_store.py**

```python
from app import store

def test_create_and_roundtrip():
    st = store.create_project("测试项目", {"voice": "zh-CN-YunxiNeural"})
    assert (store.f_raw(st.id)).exists()
    st2 = store.load_project(st.id)
    assert st2.title == "测试项目" and st2.stages["ingest"] == "pending"
    st2.stages["ingest"] = "done"; store.save_project(st2)
    assert store.load_project(st.id).stages["ingest"] == "done"

def test_write_text_atomic_and_mkdir():
    st = store.create_project("p2", {})
    p = store.write_text(store.f_script(st.id, 5), "你好。")
    assert p.endswith("slide-05.md") and store.read_text(p) == "你好。"

def test_list_projects_sorted():
    a = store.create_project("甲", {}); b = store.create_project("乙", {})
    assert [s.title for s in store.list_projects()][0] in ("甲", "乙")
    assert len(store.list_projects()) == 2
```

- [ ] **Step 2: Run** `pytest tests/test_store.py -v` → FAIL（store 不存在）
- [ ] **Step 3: 写实现**

```python
# app/store.py
import json, os, uuid, datetime
from pathlib import Path
from pydantic import TypeAdapter
from app import config
from app.models import ProjectState, new_stages

def _atomic_write(path: Path, data: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(data, encoding="utf-8")
    os.replace(tmp, path)

def _pid() -> str: return uuid.uuid4().hex[:8]

def create_project(title: str, cfg: dict) -> ProjectState:
    from app.models import ProjectConfig
    st = ProjectState(id=_pid(), title=title,
                      created_at=datetime.datetime.now().isoformat(timespec="seconds"),
                      config=ProjectConfig(**cfg), stages=new_stages(), slides=[])
    for sub in ("material/raw", "material/parsed", "material/index", "deck/png",
                "scripts", "audio/cache", "output"):
        (config.project_dir(st.id) / sub).mkdir(parents=True, exist_ok=True)
    (config.project_dir(st.id) / "deck" / "theme.css").write_text("", encoding="utf-8")
    save_project(st)
    return st

def _pjson(pid: str) -> Path: return config.project_dir(pid) / "project.json"

def save_project(st: ProjectState) -> None:
    _atomic_write(_pjson(st.id), st.model_dump_json(indent=2))

def load_project(pid: str) -> ProjectState:
    return ProjectState.model_validate_json(_pjson(pid).read_text(encoding="utf-8"))

def list_projects() -> list[ProjectState]:
    if not config.PROJECTS_DIR.exists(): return []
    out = [load_project(p.parent.name) for p in config.PROJECTS_DIR.glob("*/project.json")]
    return sorted(out, key=lambda s: s.created_at, reverse=True)

def read_text(p: Path | str) -> str:
    return Path(p).read_text(encoding="utf-8")

def write_text(p: Path | str, s: str) -> str:
    _atomic_write(Path(p), s); return str(Path(p).resolve())

def sl(n: int) -> str: return f"slide-{n:02d}"

def f_raw(pid): return config.project_dir(pid) / "material" / "raw"
def f_parsed(pid): return config.project_dir(pid) / "material" / "parsed"
def f_index(pid): return config.project_dir(pid) / "material" / "index"
def f_deck(pid): return config.project_dir(pid) / "deck"
def f_pngs(pid): return config.project_dir(pid) / "deck" / "png"
def f_outline(pid): return f_deck(pid) / "outline.json"
def f_slides_md(pid): return f_deck(pid) / "slides.md"
def f_script(pid, n): return config.project_dir(pid) / "scripts" / f"{sl(n)}.md"
def f_audio(pid, n): return config.project_dir(pid) / "audio" / f"{sl(n)}.mp3"
def f_audio_meta(pid, n): return config.project_dir(pid) / "audio" / f"{sl(n)}.json"
def f_cache(pid): return config.project_dir(pid) / "audio" / "cache"
def f_timeline(pid): return config.project_dir(pid) / "timeline.json"
def f_out(pid): return config.project_dir(pid) / "output"
```

- [ ] **Step 4: Run** `pytest tests/test_store.py -v` → PASS
- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: 项目仓库与原子读写"`

---

### Task 4: TTS 适配器与缓存（pipeline/tts.py）

**Files:**
- Create: `app/pipeline/__init__.py`, `app/pipeline/tts.py`, `tests/fakes.py`, `tests/test_tts.py`

**Interfaces:**
- Produces:
  - `cache_key(text: str, voice: str) -> str`
  - `class TTSProvider(Protocol): async def synth(self, text: str, voice: str) -> AudioResult`
  - `class EdgeTTSProvider:` 实现协议（edge_tts.Communicate 流式收集 `WordBoundary` 事件，offset/duration 为 100ns 单位 → 除以 1e7 转秒；时长用 `ffprobe_duration()` 探测）
  - `async ffprobe_duration(path: Path) -> float`（ffprobe `-show_entries format=duration`）
  - `async synth_page(provider, *, text, voice, cache_dir: Path, out_mp3: Path, out_meta: Path) -> AudioResult`：命中缓存（`cache_dir/{key}.mp3` + `.json`）直接复用；未命中则 synth→写缓存→拷贝到 out_mp3，并把 `{duration, word_boundaries}` 写 out_meta（JSON，Task 5/14 消费）

- [ ] **Step 1: 写 tests/fakes.py（本文件被 Task 11/14/17 复用）**

```python
# tests/fakes.py
from pathlib import Path
from app.models import AudioResult, WordBoundary

class FakeTTSProvider:
    """罐头音频 + 可预测的逐字边界（每字 0.5s 起、0.4s 长）；写临时 mp3 供缓存拷贝"""
    def __init__(self): self.calls: list[str] = []
    async def synth(self, text: str, voice: str) -> AudioResult:
        import tempfile
        self.calls.append(text)
        bounds = [WordBoundary(offset=0.5 * i, duration=0.4, text=ch)
                  for i, ch in enumerate(text) if ch not in "。！？；，、 "]
        tmp = Path(tempfile.mkstemp(suffix=".mp3")[1]); tmp.write_bytes(b"\xff\xfbFAKE")
        return AudioResult(path=str(tmp), duration=0.5 * len(bounds) + 0.1, word_boundaries=bounds)

class FakeStructuredLLM:
    """按 schema 类型返回罐头结构化结果的假 LLM（Task 11/14 使用）"""
    def __init__(self, responses: dict): self.responses = responses
    def with_structured_output(self, schema, **kw):
        from langchain_core.runnables import RunnableLambda
        payload = self.responses[schema.__name__]
        return RunnableLambda(lambda _msgs: payload.model_copy(deep=True))
```

- [ ] **Step 2: 写 tests/test_tts.py**

```python
import json, pytest
from pathlib import Path
from app.pipeline import tts
from tests.fakes import FakeTTSProvider

async def test_cache_key_stable():
    assert tts.cache_key("abc", "v1") == tts.cache_key("abc", "v1")
    assert tts.cache_key("abc", "v1") != tts.cache_key("abd", "v1")

async def test_synth_page_miss_then_hit(tmp_path):
    prov = FakeTTSProvider()
    out_mp3, out_meta = tmp_path / "a.mp3", tmp_path / "a.json"
    r1 = await tts.synth_page(prov, text="神经元。", voice="v", cache_dir=tmp_path/"c",
                              out_mp3=out_mp3, out_meta=out_meta)
    assert len(prov.calls) == 1 and out_mp3.exists()
    meta = json.loads(out_meta.read_text("utf-8"))
    assert meta["duration"] == r1.duration and len(meta["word_boundaries"]) == 3
    r2 = await tts.synth_page(prov, text="神经元。", voice="v", cache_dir=tmp_path/"c",
                              out_mp3=tmp_path/"b.mp3", out_meta=tmp_path/"b.json")
    assert len(prov.calls) == 1 and r2.duration == r1.duration  # 第二次命中缓存

async def test_ffprobe_duration_missing_binary(monkeypatch):
    async def boom(*a, **k): raise RuntimeError("no ffprobe")
    monkeypatch.setattr(tts, "ffprobe_duration", boom)
    with pytest.raises(RuntimeError):
        await tts.ffprobe_duration(Path("x.mp3"))
```

- [ ] **Step 3: Run** `pytest tests/test_tts.py -v` → FAIL
- [ ] **Step 4: 写实现**

```python
# app/pipeline/tts.py
import asyncio, hashlib, json, shutil
from pathlib import Path
from typing import Protocol
from app.models import AudioResult, WordBoundary

def cache_key(text: str, voice: str) -> str:
    return hashlib.sha1(f"{voice}::{text}".encode("utf-8")).hexdigest()[:16]

async def ffprobe_duration(path: Path) -> float:
    proc = await asyncio.create_subprocess_exec(
        "ffprobe", "-v", "error", "-show_entries", "format=duration",
        "-of", "default=nw=1:nk=1", str(path), stdout=asyncio.subprocess.PIPE)
    out, _ = await proc.communicate()
    return float(out.strip())

class TTSProvider(Protocol):
    async def synth(self, text: str, voice: str) -> AudioResult: ...

class EdgeTTSProvider:
    async def synth(self, text: str, voice: str) -> AudioResult:
        import edge_tts, tempfile
        tmp = Path(tempfile.mkstemp(suffix=".mp3")[1])
        bounds: list[WordBoundary] = []
        com = edge_tts.Communicate(text, voice)
        with tmp.open("wb") as f:
            async for chunk in com.stream():
                if chunk["type"] == "audio": f.write(chunk["data"])
                elif chunk["type"] == "WordBoundary":
                    bounds.append(WordBoundary(offset=chunk["offset"] / 1e7,
                                               duration=chunk["duration"] / 1e7,
                                               text=chunk["text"]))
        return AudioResult(path=str(tmp), duration=await ffprobe_duration(tmp),
                           word_boundaries=bounds)

async def synth_page(provider, *, text: str, voice: str, cache_dir: Path,
                     out_mp3: Path, out_meta: Path) -> AudioResult:
    key = cache_key(text, voice)
    cp, cj = cache_dir / f"{key}.mp3", cache_dir / f"{key}.json"
    cache_dir.mkdir(parents=True, exist_ok=True)
    if not (cp.exists() and cj.exists()):
        res = await provider.synth(text, voice)
        shutil.copyfile(res.path, cp)
        cj.write_text(json.dumps({"duration": res.duration,
                                  "word_boundaries": [b.model_dump() for b in res.word_boundaries]}),
                      encoding="utf-8")
    res = AudioResult(path=str(cp), duration=json.loads(cj.read_text("utf-8"))["duration"],
                      word_boundaries=[WordBoundary(**b) for b in
                                       json.loads(cj.read_text("utf-8"))["word_boundaries"]])
    out_mp3.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(cp, out_mp3)
    out_meta.write_text(json.dumps({"duration": res.duration,
                                    "word_boundaries": [b.model_dump() for b in res.word_boundaries]}),
                        encoding="utf-8")
    return res
```

- [ ] **Step 5: Run** `pytest tests/test_tts.py -v` → PASS（另跑真实 TTS 冒烟：`python -c "import asyncio;from app.pipeline.tts import EdgeTTSProvider,ffprobe_duration;from pathlib import Path;r=asyncio.run(EdgeTTSProvider().synth('神经元是大脑的基本单位。','zh-CN-YunxiNeural'));print(r.duration, len(r.word_boundaries))"` 预期 duration>0 且 boundaries>0 —— 需联网，失败则记录环境问题，不阻塞）
- [ ] **Step 6: Commit** `git add -A && git commit -m "feat: TTSProvider 适配器 + hash 缓存"`

---

### Task 5: 时间轴与字幕（pipeline/timeline.py，含句切分）

**Files:**
- Create: `app/pipeline/timeline.py`, `tests/test_timeline.py`

**Interfaces:**
- Produces:
  - `split_sentences(text: str) -> list[str]`（按 `。！？；` 切，过滤纯空白）
  - `align_sentences(text: str, bounds: list[WordBoundary]) -> list[SubtitleCue]`（字级→句级，页内相对时间；无边界命中的句并入前一句，首句无命中则丢弃）
  - `build_timeline(slides: list[SlideMeta], meta_of: dict[int, dict], gap: float, resolution: list[int]) -> Timeline`（`meta_of[n]` = Task 4 写的 out_meta JSON 内容；segment.duration = audio_duration + gap；cue 全局时间 = 段起始 + 页内相对）
  - `write_ass(tl: Timeline, path: Path) -> None`；`write_srt(tl: Timeline, path: Path) -> None`
  - `fmt_ass(t: float) -> str`（`H:MM:SS.cc`）、`fmt_srt(t: float) -> str`（`HH:MM:SS,mmm`）

- [ ] **Step 1: 写测试 tests/test_timeline.py**

```python
from pathlib import Path
from app.models import SlideMeta, WordBoundary, SubtitleCue
from app.pipeline import timeline as T

def test_split_sentences():
    assert T.split_sentences("神经元是大脑的基本单位。它由胞体、树突和轴突构成。") == \
        ["神经元是大脑的基本单位。", "它由胞体、树突和轴突构成。"]
    assert T.split_sentences("没有结尾") == ["没有结尾"]
    assert T.split_sentences("。！？；") == []

def test_align_sentences_two_cues():
    text = "神经元是大脑的基本单位。它由胞体、树突和轴突构成。"
    bounds = [WordBoundary(offset=0.5 * i, duration=0.4, text=ch)
              for i, ch in enumerate(text) if ch not in "。，、"]
    cues = T.align_sentences(text, bounds)
    assert len(cues) == 2
    assert cues[0].text == "神经元是大脑的基本单位。" and cues[0].start == 0.0
    assert cues[1].start == 0.5 * 10  # 第二句从第11个有效字开始
    assert cues[1].end > cues[1].start

def test_align_no_bound_first_sentence_dropped():
    cues = T.align_sentences("甲。乙。", [WordBoundary(offset=0, duration=1, text="乙")])
    assert [c.text for c in cues] == ["乙。"]

def test_build_timeline_and_files(tmp_path):
    meta = {1: {"duration": 3.0, "word_boundaries": [{"offset": 0.0, "duration": 1.0, "text": "一"}]},
            2: {"duration": 2.0, "word_boundaries": [{"offset": 0.0, "duration": 1.0, "text": "二"}]}}
    slides = [SlideMeta(index=1, png="p1", audio="a1", audio_duration=3.0),
              SlideMeta(index=2, png="p2", audio="a2", audio_duration=2.0)]
    tl = T.build_timeline(slides, meta, gap=0.4, resolution=[1920, 1080])
    assert [s.duration for s in tl.segments] == [3.4, 2.4]
    assert tl.segments[1].subtitle_cues[0].start == 3.4  # 全局时间累加
    T.write_ass(tl, tmp_path / "s.ass"); T.write_srt(tl, tmp_path / "s.srt")
    assert "Dialogue:" in (tmp_path / "s.ass").read_text("utf-8")
    assert "-->" in (tmp_path / "s.srt").read_text("utf-8")

def test_fmt():
    assert T.fmt_ass(3671.25) == "1:01:11.25"
    assert T.fmt_srt(3671.25) == "01:01:11,250"
```

- [ ] **Step 2: Run** → FAIL
- [ ] **Step 3: 写实现**

```python
# app/pipeline/timeline.py
import re
from pathlib import Path
from app.models import SlideMeta, SubtitleCue, Timeline, Segment, WordBoundary

_SENT_RE = re.compile(r"[^。！？；]*[。！？；]|[^。！？；]+$")
_SKIP = set("。！？；，、：；()（）\"'“”‘’…—·,.!?;: \t\n\r")

def split_sentences(text: str) -> list[str]:
    return [m.strip() for m in _SENT_RE.findall(text) if m.strip()]

def align_sentences(text: str, bounds: list[WordBoundary]) -> list[SubtitleCue]:
    spans: dict[int, tuple[float, float]] = {}   # 正文字符下标 -> (start,end)
    p = 0
    for b in bounds:
        q = p
        for ch in b.text:
            while q < len(text) and text[q] in _SKIP: q += 1
            if q < len(text) and text[q] == ch:
                spans[q] = (b.offset, b.offset + b.duration); q += 1
        p = q  # 容忍 TTS 归一化造成的个别失配：指针不回退
    out: list[SubtitleCue] = []
    for sent in split_sentences(text):
        start_i = text.find(sent, pos); end_i = start_i + len(sent); pos = end_i
        ss = [spans[i] for i in range(start_i, end_i) if i in spans]
        if not ss:
            if out: out[-1].text += sent  # 无边界命中：并入前一句
            continue
        out.append(SubtitleCue(start=min(a for a, _ in ss), end=max(b for _, b in ss), text=sent))
    return out

def build_timeline(slides, meta_of, gap: float, resolution: list[int]) -> Timeline:
    segs: list[Segment] = []; t = 0.0
    for s in slides:
        meta = meta_of[s.index]
        bounds = [WordBoundary(**x) for x in meta["word_boundaries"]]
        script = Path(s.script).read_text(encoding="utf-8") if s.script else ""
        cues = [SubtitleCue(start=t + c.start, end=t + c.end, text=c.text)
                for c in align_sentences(script, bounds)]
        segs.append(Segment(index=s.index, png=s.png, audio=s.audio,
                            duration=(s.audio_duration or meta["duration"]) + gap,
                            subtitle_cues=cues))
        t += segs[-1].duration
    return Timeline(resolution=resolution, segments=segs)
```

```python
def fmt_ass(t: float) -> str:
    h = int(t // 3600); m = int(t % 3600 // 60); s = t % 60
    return f"{h}:{m:02d}:{s:05.2f}"

def fmt_srt(t: float) -> str:
    h = int(t // 3600); m = int(t % 3600 // 60); s = int(t % 60); ms = round((t - int(t)) * 1000)
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"

_ASS_HEADER = """[Script Info]
ScriptType: v4.00+
PlayResX: {rx}
PlayResY: {ry}

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, OutlineColour, BackColour, Bold, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, BorderStyle
Style: Default,Microsoft YaHei,54,&H00FFFFFF,&H00000000,&H7F000000,1,2,0,2,60,60,50,1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
"""

def write_ass(tl: Timeline, path: Path) -> None:
    lines = [_ASS_HEADER.format(rx=tl.resolution[0], ry=tl.resolution[1])]
    for seg in tl.segments:
        for c in seg.subtitle_cues:
            lines.append(f"Dialogue: 0,{fmt_ass(c.start)},{fmt_ass(c.end)},Default,,0,0,0,,{c.text}")
    path.write_text("\n".join(lines), encoding="utf-8")

def write_srt(tl: Timeline, path: Path) -> None:
    blocks = []
    for i, seg in enumerate(tl.segments, 1):
        for c in seg.subtitle_cues:
            blocks.append(f"{i}\n{fmt_srt(c.start)} --> {fmt_srt(c.end)}\n{c.text}\n")
    path.write_text("\n".join(blocks), encoding="utf-8")
```

（注：ASS/SRT 序号按 cue 全局递增，测试断言只需包含 `Dialogue:`/`-->` 即可。）

- [ ] **Step 4: Run** `pytest tests/test_timeline.py -v` → PASS
- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: 句切分/字级对齐/timeline/ASS/SRT"`

---

### Task 6: ffmpeg 渲染器（pipeline/render.py）

**Files:**
- Create: `app/pipeline/render.py`, `tests/test_render.py`

**Interfaces:**
- Consumes: `Timeline/Segment`（Task 2）、`ffprobe_duration`（Task 4）
- Produces: `async render_video(tl: Timeline, base_dir: Path, work_dir: Path, final: Path, on_progress=None) -> Path`
  - `seg.png`/`seg.audio` 是相对 `base_dir` 的路径；`subtitles.ass` 必须已存在于 `work_dir`（Task 14 负责先写好）；三步：逐页分段 → concat 流复制 → 一遍编码烧 ass（ass 滤镜用 `cwd=work_dir` 相对路径避免转义）

- [ ] **Step 1: 写测试**（无 ffmpeg 则 skip）

```python
import shutil
from pathlib import Path
import pytest
from app.models import Timeline, Segment, SubtitleCue
from app.pipeline import render as R

pytestmark = pytest.mark.skipif(shutil.which("ffmpeg") is None, reason="no ffmpeg")

async def _tone(path: Path, sec: float):
    await R._run(["ffmpeg","-y","-f","lavfi","-i",
                  "sine=frequency=440:duration=%s" % sec, str(path)])

_ASS_MIN = ("[Script Info]\nPlayResX: 320\nPlayResY: 180\n\n[V4+ Styles]\n"
    "Format: Name, Fontname, Fontsize, PrimaryColour, OutlineColour, BackColour, Bold, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, BorderStyle\n"
    "Style: Default,Arial,20,&H00FFFFFF,&H00000000,&H7F000000,1,1,0,2,10,10,10,1\n\n"
    "[Events]\nFormat: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text\n"
    "Dialogue: 0,0:00:00.00,0:00:00.40,Default,,0,0,0,,你好。\n")

async def test_render_two_segments(tmp_path):
    base = tmp_path; work = base / "out" / "_work"; work.mkdir(parents=True)
    a1, a2 = base/"a1.wav", base/"a2.wav"
    await _tone(a1, 0.5); await _tone(a2, 0.5)
    tl = Timeline(resolution=[320,180], segments=[
        Segment(index=1, png="x.png", audio="a1.wav", duration=0.6,
                subtitle_cues=[SubtitleCue(start=0.0, end=0.4, text="你好。")]),
        Segment(index=2, png="x.png", audio="a2.wav", duration=0.6,
                subtitle_cues=[SubtitleCue(start=0.6, end=1.0, text="世界。")])])
    from PIL import Image
    Image.new("RGB", (320, 180)).save(base / "x.png")  # 真 PNG，ffmpeg 可解码
    (work / "subtitles.ass").write_text(_ASS_MIN, encoding="utf-8")
    final = await R.render_video(tl, base_dir=base, work_dir=work, final=base/"out"/"final.mp4")
    assert final.exists() and final.stat().st_size > 0
    dur = await R.ffprobe_duration(final)
    assert 1.0 <= dur <= 1.6
```

注：若本机 ffmpeg 对合成分段有额外约束，可把 `-framerate` 提到 25。测试中两段各 0.6s、总时长约 1.2s。

- [ ] **Step 2: Run** → FAIL
- [ ] **Step 3: 写实现**

```python
# app/pipeline/render.py
import asyncio
from pathlib import Path
from app.models import Timeline
from app.pipeline.tts import ffprobe_duration  # re-export 供测试

async def _run(cmd: list[str], cwd: Path | None = None) -> None:
    proc = await asyncio.create_subprocess_exec(
        *cmd, cwd=cwd, stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE)
    _, err = await proc.communicate()
    if proc.returncode != 0:
        raise RuntimeError(f"ffmpeg failed: {err.decode(errors='ignore')[-500:]}")

async def render_video(tl: Timeline, base_dir: Path, work_dir: Path, final: Path,
                       on_progress=None) -> Path:
    work_dir.mkdir(parents=True, exist_ok=True)
    segs = []
    for seg in tl.segments:
        out = work_dir / f"seg-{seg.index:02d}.mp4"
        await _run(["ffmpeg","-y","-loop","1","-framerate","30",
                    "-i", str(base_dir / seg.png), "-i", str(base_dir / seg.audio),
                    "-t", f"{seg.duration:.3f}", "-pix_fmt","yuv420p",
                    "-c:v","libx264","-preset","veryfast","-c:a","aac","-shortest", str(out)])
        segs.append(out)
        if on_progress: on_progress(seg.index, len(tl.segments))
    lst = work_dir / "list.txt"
    lst.write_text("".join(f"file '{s.as_posix()}'\n" for s in segs), encoding="utf-8")
    concat = work_dir / "concat.mp4"
    await _run(["ffmpeg","-y","-f","concat","-safe","0","-i",str(lst),"-c","copy",str(concat)])
    final.parent.mkdir(parents=True, exist_ok=True)
    await _run(["ffmpeg","-y","-i",str(concat),"-vf",f"ass={(work_dir/'subtitles.ass').name}",
                "-c:a","copy",str(final)], cwd=work_dir)
    return final
```

- [ ] **Step 4: Run** `pytest tests/test_render.py -v` → PASS（或 skip）
- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: ffmpeg 分段-拼接-烧字幕渲染器"`

---

### Task 7: Marp deck 组装与截图（pipeline/deck.py + theme）

**Files:**
- Create: `app/pipeline/deck.py`, `tests/test_deck.py`

**Interfaces:**
- Produces: `DEFAULT_THEME: str`（声明 `section{width:1920px;height:1080px}`，含 cover/figure 两个版式 class）、`init_deck(deck_dir: Path) -> None`（写 theme.css）、`assemble(parts: list[str]) -> str`（用 `\n\n---\n\n` 连接）、`render_pngs(deck_dir: Path) -> list[Path]`（npx marp 截图 → `png/slide-NN.png`；产物为空则 RuntimeError）

- [ ] **Step 1: 写测试**

```python
import shutil
import pytest
from app.pipeline import deck as D

def test_assemble_and_theme(tmp_path):
    assert D.assemble(["# A\n- 1", "# B"]) == "# A\n- 1\n\n---\n\n# B"
    D.init_deck(tmp_path)
    css = (tmp_path / "theme.css").read_text("utf-8")
    assert "1920px" in css and "1080px" in css and "@theme lessonforge" in css

@pytest.mark.skipif(shutil.which("npx") is None, reason="no npx")
def test_render_pngs_real(tmp_path):
    D.init_deck(tmp_path)
    (tmp_path / "slides.md").write_text("# 第一页\n\n---\n\n# 第二页", encoding="utf-8")
    pngs = D.render_pngs(tmp_path)
    assert [p.name for p in pngs] == ["slide-01.png", "slide-02.png"]
    from PIL import Image
    assert Image.open(pngs[0]).size == (1920, 1080)
```

- [ ] **Step 2: Run** → FAIL
- [ ] **Step 3: 写实现**

```python
# app/pipeline/deck.py
import shutil, subprocess
from pathlib import Path

DEFAULT_THEME = '''/* @theme lessonforge */
section {
  width: 1920px; height: 1080px; padding: 96px 120px;
  font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
  background: #0f172a; color: #e2e8f0;
}
h1 { color: #38bdf8; font-size: 72px; }
h2 { color: #7dd3fc; font-size: 56px; }
ul li { font-size: 40px; margin: 16px 0; }
section.cover { text-align: center; }
section.cover h1 { font-size: 96px; margin-top: 320px; }
section.figure { text-align: center; }
section.figure img { max-width: 1500px; max-height: 800px; }
'''

def init_deck(deck_dir: Path) -> None:
    deck_dir.mkdir(parents=True, exist_ok=True)
    (deck_dir / "theme.css").write_text(DEFAULT_THEME, encoding="utf-8")

def assemble(parts: list[str]) -> str:
    return "\n\n---\n\n".join(p.strip() for p in parts)

def render_pngs(deck_dir: Path) -> list[Path]:
    npx = shutil.which("npx")
    if not npx: raise RuntimeError("未找到 npx（需要 Node.js）")
    out = deck_dir / "png"; out.mkdir(exist_ok=True)
    tmp = deck_dir / "_marp_tmp"
    r = subprocess.run([npx, "marp", str(deck_dir / "slides.md"),
                        "--theme", str(deck_dir / "theme.css"),
                        "--images", "png", "--allow-local-files", "-o", str(tmp)],
                       capture_output=True, timeout=600)
    pngs = sorted(tmp.glob("*.png")) if tmp.exists() else []
    if r.returncode != 0 or not pngs:
        raise RuntimeError(f"marp 截图失败: {r.stderr.decode(errors='ignore')[-300:]}")
    res = []
    for i, p in enumerate(pngs, 1):
        dst = out / f"slide-{i:02d}.png"
        shutil.move(str(p), dst); res.append(dst)
    shutil.rmtree(tmp, ignore_errors=True)
    return res
```

- [ ] **Step 4: Run** → PASS（无 npx 时真实截图用例 skip）
- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: Marp 主题/组装/截图"`

---

### Task 8: 摄取——PyMuPDFParser（含 LLM 章节兜底）

**Files:**
- Create: `app/pipeline/ingest.py`, `tests/test_ingest.py`

**Interfaces:**
- Consumes: `Chapter`（Task 2）、`ask_structured`（本任务在 chains.py 建立；Task 11 在同文件追加三条链）
- Produces:
  - `make_parser(cfg, llm=None) -> MaterialParser`：`cfg.parser == pymupdf` → `PyMuPDFParser(llm)`；否则 → Task 9 `MinerUParser(mode=cfg.parser.value)`
  - `class PyMuPDFParser(llm=None)`；`parse(pdf: Path, out_dir: Path) -> list[Chapter]`
    - 有书签（TOC level-1）→ 按 `[(title, start_page, end_page)]` 切章
    - 无书签且有 llm → `_llm_spans(doc)`：每页前 80 字喂 schema `ChapterStarts{chapters:[{title,start_page}]}`（定义在本文件，pydantic）；非法/失败 → 空列表退化为单章
    - 都没有 → 单章（pdf 文件名做标题）
    - 章节文本 = 各页 `get_text()` 拼接，头部 `# 标题`；抽取 ≥200px 图片到 `out_dir/images/`，文件名记入 `Chapter.images`；每章 md 写盘 `{i:02d}-{安全标题}.md`

- [ ] **Step 1: 写测试**（fixture PDF 用 fitz 现做）

```python
import fitz
from pathlib import Path
from app.pipeline.ingest import PyMuPDFParser

def _make_pdf(path: Path, with_toc: bool):
    doc = fitz.open()
    for i, txt in enumerate(["神经元绪论", "突触传递", "附录参考"], 1):
        p = doc.new_page(); p.insert_text((72, 100), f"{txt} p{i} 正文内容")
    if with_toc: doc.set_toc([[1, "第一章", 1], [1, "第二章", 2]])
    doc.save(path); doc.close()

def test_toc_split(tmp_path):
    pdf = tmp_path / "book.pdf"; _make_pdf(pdf, True)
    chs = PyMuPDFParser().parse(pdf, tmp_path / "parsed")
    assert [c.title for c in chs] == ["第一章", "第二章"]
    assert chs[0].markdown.startswith("# 第一章") and "正文内容" in chs[0].markdown
    assert (tmp_path / "parsed" / "01-第一章.md").exists()

def test_no_toc_no_llm_single_chapter(tmp_path):
    pdf = tmp_path / "book.pdf"; _make_pdf(pdf, False)
    chs = PyMuPDFParser().parse(pdf, tmp_path / "parsed")
    assert len(chs) == 1 and chs[0].title == "book"

def test_no_toc_llm_fallback(tmp_path):
    from tests.fakes import FakeStructuredLLM
    from app.pipeline import ingest as I
    fake = FakeStructuredLLM({"ChapterStarts": I.ChapterStarts(chapters=[
        I.ChapterStart(title="绪论", start_page=1),
        I.ChapterStart(title="突触", start_page=2)])})
    pdf = tmp_path / "book.pdf"; _make_pdf(pdf, False)
    chs = PyMuPDFParser(llm=fake).parse(pdf, tmp_path / "parsed")
    assert [c.title for c in chs] == ["绪论", "突触"]
```

- [ ] **Step 2: Run** → FAIL
- [ ] **Step 3: 写实现**

```python
# app/pipeline/ingest.py
import re
from pathlib import Path
import fitz
from pydantic import BaseModel
from app.models import Chapter, ParserKind

class ChapterStart(BaseModel): title: str; start_page: int
class ChapterStarts(BaseModel): chapters: list[ChapterStart]

class MaterialParser:
    def parse(self, pdf: Path, out_dir: Path) -> list[Chapter]: raise NotImplementedError

def _safe(name: str) -> str:
    return re.sub(r"[\\/:*?\"<>|\s]", "_", name)[:40] or "untitled"

def _toc_spans(doc) -> list[tuple[str, int, int]]:
    toc = [t for t in doc.get_toc() if t[0] == 1]
    if not toc: return []
    spans = []
    for i, (_lvl, title, pg) in enumerate(toc):
        end = toc[i + 1][2] - 1 if i + 1 < len(toc) else doc.page_count
        spans.append((title.strip(), pg, max(end, pg)))
    return spans

def _save_images(page, img_dir: Path, min_px: int = 200) -> list[str]:
    rels = []
    for img in page.get_images(full=True):
        pix = fitz.Pixmap(page.parent, img[0])
        if pix.width < min_px or pix.height < min_px: continue
        if pix.colorspace and pix.colorspace.n > 3: pix = fitz.Pixmap(fitz.csRGB, pix)
        name = f"img-p{page.number + 1}-{img[0]}.png"
        pix.save(img_dir / name); rels.append(name)
    return rels

class PyMuPDFParser(MaterialParser):
    def __init__(self, llm=None): self.llm = llm
    def parse(self, pdf: Path, out_dir: Path) -> list[Chapter]:
        doc = fitz.open(pdf)
        (out_dir / "images").mkdir(parents=True, exist_ok=True)
        spans = _toc_spans(doc)
        if not spans and self.llm is not None: spans = self._llm_spans(doc)
        if not spans: spans = [(pdf.stem, 1, doc.page_count)]
        chapters = []
        for i, (title, a, b) in enumerate(spans, 1):
            imgs: list[str] = []; parts: list[str] = []
            for pno in range(a - 1, b):
                page = doc[pno]
                parts.append(page.get_text())
                imgs += _save_images(page, out_dir / "images")
            md = f"# {title}\n\n" + "\n".join(parts)
            (out_dir / f"{i:02d}-{_safe(title)}.md").write_text(md, encoding="utf-8")
            chapters.append(Chapter(title=title, markdown=md, images=imgs))
        return chapters
    def _llm_spans(self, doc) -> list[tuple[str, int, int]]:
        from app.pipeline.chains import ask_structured   # 延迟导入避免循环
        from langchain_core.messages import HumanMessage, SystemMessage
        sample = "\n".join(f"第{i+1}页：{doc[i].get_text()[:80]}" for i in range(doc.page_count))
        msgs = [SystemMessage(content="从教材页面采样中识别章节起始页，start_page 为 1 起始页码。"),
                HumanMessage(content=sample)]
        res, err = ask_structured(self.llm, ChapterStarts, msgs)
        if err or not res or not res.chapters: return []
        starts = sorted({(c.title.strip(), c.start_page) for c in res.chapters
                         if 1 <= c.start_page <= doc.page_count})
        return [(t, pg, (starts[j+1][1] - 1 if j + 1 < len(starts) else doc.page_count))
                for j, (t, pg) in enumerate(starts)]

def make_parser(cfg, llm=None) -> MaterialParser:
    if cfg.parser == ParserKind.pymupdf: return PyMuPDFParser(llm=llm)
    from app.pipeline.mineru import MinerUParser
    return MinerUParser(mode=cfg.parser.value)
```

同时新建 `app/pipeline/chains.py`（本任务只含 `ask_structured`；Task 11 追加三条链）：

```python
# app/pipeline/chains.py
from langchain_core.messages import HumanMessage

def ask_structured(llm, schema, messages: list, max_retries: int = 2):
    """返回 (schema 实例 | None, 最后一次异常 | None)。失败把报错摘要追加进消息重试。"""
    msgs = list(messages)
    last_err: Exception | None = None
    for _ in range(max_retries + 1):
        try:
            structured = llm.with_structured_output(schema)
            res = structured.invoke(msgs)
            if isinstance(res, schema): return res, None
            raise TypeError(f"期望 {schema.__name__}，得到 {type(res).__name__}")
        except Exception as e:                 # noqa: BLE001 —— 统一捕获后重试
            last_err = e
            msgs = msgs + [HumanMessage(content=f"上次输出不合法（{str(e)[:200]}），请严格按 schema 重新输出。")]
    return None, last_err
```

- [ ] **Step 4: Run** `pytest tests/test_ingest.py -v` → PASS（chains.ask_structured 已就位）
- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: PyMuPDF 摄取与章节切分"`

---

### Task 9: 摄取——MinerUParser（local / API）

**Files:**
- Create: `app/pipeline/mineru.py`, `tests/test_mineru.py`

**Interfaces:**
- Produces: `class MinerUParser(MaterialParser): def __init__(self, mode: str)`；`parse(pdf, out_dir) -> list[Chapter]`
  - `mineru_local`：`subprocess.run(["mineru", "-p", pdf, "-o", tmp], check=True, timeout=3600)` → `tmp.rglob("*.md")` 取第一个
  - `mineru_api`：`httpx.post(MINERU_API_URL, files={"file": ...}, Bearer MINERU_API_TOKEN)`（两环境变量，doctor 检查）→ 响应 JSON 递归找 `.zip` 链接 → 下载解压 → 同样找 `*.md`
  - 公共 `_md_to_chapters(md_path, out_dir) -> list[Chapter]`：按 `(?m)^# ` 切章；`images/` 复制到 `out_dir/images/`；每章 md 写盘

- [ ] **Step 1: 写测试**

```python
from pathlib import Path
from app.pipeline.mineru import MinerUParser, _md_to_chapters

MD = '''# 第一章 神经元

正文甲 ![图](images/a.png)

# 第二章 突触

正文乙
'''

def test_md_to_chapters(tmp_path):
    src = tmp_path / "auto"; (src / "images").mkdir(parents=True)
    (src / "out.md").write_text(MD, encoding="utf-8")
    (src / "images" / "a.png").write_bytes(b"png")
    chs = _md_to_chapters(src / "out.md", tmp_path / "parsed")
    assert [c.title for c in chs] == ["第一章 神经元", "第二章 突触"]
    assert chs[0].images == ["a.png"]
    assert (tmp_path / "parsed" / "images" / "a.png").exists()

def test_local_mode_via_fake_subprocess(tmp_path, monkeypatch):
    def fake_run(cmd, **kw):
        out = Path(cmd[cmd.index("-o") + 1]); d = out / "book" / "auto"
        d.mkdir(parents=True); (d / "book.md").write_text(MD, encoding="utf-8")
        (d / "images").mkdir(); (d / "images" / "a.png").write_bytes(b"png")
    monkeypatch.setattr("app.pipeline.mineru.subprocess.run", fake_run)
    pdf = tmp_path / "book.pdf"; pdf.write_bytes(b"%PDF-fake")
    chs = MinerUParser(mode="mineru_local").parse(pdf, tmp_path / "parsed")
    assert len(chs) == 2
```

- [ ] **Step 2: Run** → FAIL
- [ ] **Step 3: 写实现**

```python
# app/pipeline/mineru.py
import os, re, shutil, subprocess, tempfile, zipfile
from pathlib import Path
import httpx
from app.models import Chapter
from app.pipeline.ingest import MaterialParser, _safe

def _find_zip_url(obj):
    if isinstance(obj, str) and obj.endswith(".zip"): return obj
    if isinstance(obj, dict):
        for v in obj.values():
            r = _find_zip_url(v)
            if r: return r
    if isinstance(obj, list):
        for v in obj:
            r = _find_zip_url(v)
            if r: return r
    return None

def _md_to_chapters(md_path: Path, out_dir: Path) -> list[Chapter]:
    content = md_path.read_text(encoding="utf-8")
    img_src = md_path.parent / "images"
    (out_dir / "images").mkdir(parents=True, exist_ok=True)
    if img_src.exists():
        for f in img_src.iterdir(): shutil.copy(f, out_dir / "images" / f.name)
    chapters = []
    for i, body in enumerate(p for p in re.split(r"(?m)^# ", content) if p.strip()):
        lines = body.splitlines()
        title = lines[0].strip() or f"第{i}章"
        md = f"# {title}\n" + "\n".join(lines[1:])
        (out_dir / f"{i:02d}-{_safe(title)}.md").write_text(md, encoding="utf-8")
        imgs = re.findall(r"images/([\w.-]+)", body)
        chapters.append(Chapter(title=title, markdown=md,
                                images=[x for x in imgs if (out_dir / "images" / x).exists()]))
    return chapters

class MinerUParser(MaterialParser):
    def __init__(self, mode: str):
        self.mode = mode
        self.api_url = os.getenv("MINERU_API_URL", "")
        self.api_token = os.getenv("MINERU_API_TOKEN", "")
    def parse(self, pdf: Path, out_dir: Path) -> list[Chapter]:
        tmp = Path(tempfile.mkdtemp(prefix="mineru-"))
        if self.mode == "mineru_local":
            subprocess.run(["mineru", "-p", str(pdf), "-o", str(tmp)],
                           check=True, capture_output=True, timeout=3600)
        else:
            self._api_parse(pdf, tmp)
        mds = sorted(tmp.rglob("*.md"))
        if not mds: raise RuntimeError("MinerU 产物中未找到 markdown")
        return _md_to_chapters(mds[0], out_dir)
    def _api_parse(self, pdf: Path, tmp: Path) -> None:
        if not self.api_url: raise RuntimeError("未配置 MINERU_API_URL")
        with pdf.open("rb") as f:
            r = httpx.post(self.api_url, files={"file": (pdf.name, f)},
                           headers={"Authorization": f"Bearer {self.api_token}"}, timeout=600)
        r.raise_for_status()
        zip_url = _find_zip_url(r.json())
        if not zip_url: raise RuntimeError(f"MinerU API 响应中无 zip: {str(r.json())[:200]}")
        zpath = tmp / "result.zip"
        zpath.write_bytes(httpx.get(zip_url, timeout=600).content)
        with zipfile.ZipFile(zpath) as z: z.extractall(tmp)
```

- [ ] **Step 4: Run** → PASS
- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: MinerU 摄取（local/API 双模式）"`

---

### Task 10: 向量索引与检索（pipeline/index.py）

**Files:**
- Create: `app/pipeline/index.py`, `tests/test_index.py`

**Interfaces:**
- Consumes: `Chapter`、`ProjectConfig.embedding`
- Produces: `make_embeddings(cfg: ProjectConfig) -> OpenAIEmbeddings`（base_url/api_key/model 取自 `cfg.embedding`，key 空则 "EMPTY"）、`build_index(chapters, embeddings, persist_dir) -> None`（RecursiveCharacterTextSplitter 1200/150，metadata 带 chapter）、`retrieve(query, embeddings, persist_dir, k=4) -> list[str]`（空库/不存在返回 []）

- [ ] **Step 1: 写测试**

```python
from langchain_core.embeddings.fake import DeterministicFakeEmbedding
from app.models import Chapter
from app.pipeline.index import build_index, retrieve

def test_build_and_retrieve(tmp_path):
    emb = DeterministicFakeEmbedding(size=64)
    chs = [Chapter(title="章一", markdown="神经元是大脑的基本单位。" * 50),
           Chapter(title="章二", markdown="突触分为电突触和化学突触。" * 50)]
    build_index(chs, emb, tmp_path)
    hits = retrieve("突触 类型", emb, tmp_path, k=2)
    assert len(hits) == 2 and all(isinstance(h, str) for h in hits)

def test_empty_index(tmp_path):
    emb = DeterministicFakeEmbedding(size=64)
    assert retrieve("任意", emb, tmp_path / "none", k=4) == []
```

- [ ] **Step 2: Run** → FAIL
- [ ] **Step 3: 写实现**

```python
# app/pipeline/index.py
from pathlib import Path
from langchain_chroma import Chroma
from langchain_openai import OpenAIEmbeddings
from langchain_text_splitters import RecursiveCharacterTextSplitter
from app.models import Chapter, ProjectConfig

def make_embeddings(cfg: ProjectConfig) -> OpenAIEmbeddings:
    e = cfg.embedding
    return OpenAIEmbeddings(base_url=e.base_url, api_key=e.api_key or "EMPTY", model=e.model)

def build_index(chapters: list[Chapter], embeddings, persist_dir: Path) -> None:
    splitter = RecursiveCharacterTextSplitter(chunk_size=1200, chunk_overlap=150)
    texts: list[str] = []; metas: list[dict] = []
    for ch in chapters:
        for ck in splitter.split_text(ch.markdown):
            texts.append(ck); metas.append({"chapter": ch.title})
    if texts:
        Chroma.from_texts(texts, embedding=embeddings,
                          persist_directory=str(persist_dir), metadatas=metas)

def retrieve(query: str, embeddings, persist_dir: Path, k: int = 4) -> list[str]:
    if not Path(persist_dir).exists(): return []
    db = Chroma(persist_directory=str(persist_dir), embedding_function=embeddings)
    if db._collection.count() == 0: return []
    return [d.page_content for d in db.similarity_search(query, k=k)]
```

- [ ] **Step 4: Run** → PASS
- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: Chroma 索引与检索"`

---

### Task 11: LLM 工厂与 LCEL 链（pipeline/llm.py + chains.py）

**Files:**
- Create: `app/pipeline/llm.py`, `tests/test_chains.py`
- Modify: `app/pipeline/chains.py`（Task 8 已建 `ask_structured`，本任务追加三条链）、`tests/fakes.py`（重写 `FakeStructuredLLM`：消息捕获 + `invoke` 支持；新增 `FlakyStructuredLLM`、`RawOnlyLLM`）

**Interfaces:**
- Produces:
  - `make_chat_model(cfg: LLMConfig) -> ChatOpenAI`（temperature=0.3，api_key 空则 "EMPTY"）
  - `ask_structured(llm, schema, messages, max_retries=2) -> tuple[Any|None, Exception|None]`（失败把异常摘要追加进消息重试；Task 8 已引用）
  - `run_outline(llm, chapter: str, max_slides: int = 12) -> Outline`（PromptTemplate -> with_structured_output(Outline)）
  - `run_slide(llm, slide: OutlineSlide, context: str) -> str`（产出单页 Marp markdown；Prompt 含“数字用中文汉字”约束）
  - `run_script(llm, slide: OutlineSlide, slide_md: str, context: str) -> ScriptPart`（讲稿 150-300 字/页，检索增强 context 由 orchestrator 提供）

- [ ] **Step 1: 写测试**

```python
from app.models import Outline, OutlineSlide, ScriptPart
from app.pipeline.chains import run_outline, run_slide, run_script, ask_structured
from tests.fakes import FakeStructuredLLM, FlakyStructuredLLM, RawOnlyLLM
from langchain_core.messages import HumanMessage

OUTLINE = Outline(slides=[OutlineSlide(title="神经元", points=["定义", "结构"]),
                          OutlineSlide(title="突触", points=["分类"])])

def test_outline_chain():
    fake = FakeStructuredLLM({"Outline": OUTLINE})
    o = run_outline(fake, "章一内容", max_slides=5)
    assert len(o.slides) == 2 and o.slides[0].title == "神经元"
    assert "章一内容" in fake.last_messages[-1].content

def test_slide_chain_mentions_numbers_rule():
    fake = FakeStructuredLLM({})
    md = run_slide(fake, OUTLINE.slides[0], "ctx")
    assert md.startswith("# ")
    assert "中文" in fake.last_messages[0].content  # 汉字数字约束在 system prompt

def test_script_chain():
    fake = FakeStructuredLLM({"ScriptPart": ScriptPart(narration="神经元，是大脑的基本单位。" * 20)})
    s = run_script(fake, OUTLINE.slides[0], "# 神经元", "ctx")
    assert s.narration.startswith("神经元")

def test_ask_structured_retries_then_recovers():
    flaky = FlakyStructuredLLM({"Outline": OUTLINE}, fail_times=2)
    res, err = ask_structured(flaky, Outline, [HumanMessage(content="x")])
    assert res is not None and err is None and flaky.calls == 3

def test_ask_structured_unknown_schema_returns_err():
    fake = FakeStructuredLLM({})                       # 无对应罐头 → 每次都失败
    res, err = ask_structured(fake, Outline, [HumanMessage(content="x")])
    assert res is None and err is not None

def test_ask_structured_raw_fallback():
    raw = RawOnlyLLM("标题：封面\n要点：定义")
    res, err = ask_structured(raw, Outline, [HumanMessage(content="x")])
    assert res is None and err is not None
```

- [ ] **Step 2: Run** → FAIL

- [ ] **Step 3: 写实现**

```python
# app/pipeline/llm.py
from langchain_openai import ChatOpenAI
from app.models import LLMConfig

def make_chat_model(cfg: LLMConfig) -> ChatOpenAI:
    return ChatOpenAI(base_url=cfg.base_url, api_key=cfg.api_key or "EMPTY",
                      model=cfg.model, temperature=0.3, timeout=120)
```

```python
# app/pipeline/chains.py（追加；ask_structured 已在 Task 8 建立）
from langchain_core.prompts import ChatPromptTemplate
from app.models import Outline, OutlineSlide, ScriptPart

STRUCT_SYS = "你是严谨的教材讲师。只依据给定材料输出，不得编造事实；数字一律写中文汉字（如“一百二十”），避免 TTS 误读。"

def run_outline(llm, chapter: str, max_slides: int = 12) -> Outline:
    prompt = ChatPromptTemplate.from_messages([
        ("system", STRUCT_SYS + " 阅读教材章节，规划不超过 {max_slides} 页的讲课大纲，第一页为封面。"),
        ("human", "章节内容：\n{chapter}"),
    ])
    msgs = prompt.format_messages(max_slides=max_slides, chapter=chapter[:8000])
    res, err = ask_structured(llm, Outline, msgs)
    if err or res is None: raise RuntimeError(f"outline 生成失败: {err}")
    return res

SLIDE_TPL = ("为以下大纲页写单页 Marp markdown（不要 front-matter，不要 --- 分隔符）。"
             "标题一行；要点 3-5 条、每条不超过 20 字。\n"
             "大纲页：{page_json}\n教材片段：{context}")

def run_slide(llm, slide: OutlineSlide, context: str) -> str:
    prompt = ChatPromptTemplate.from_messages([
        ("system", STRUCT_SYS), ("human", SLIDE_TPL)])
    msgs = prompt.format_messages(page_json=slide.model_dump_json(),
                                  context=context[:3000])
    raw = llm.invoke(msgs).content
    md = raw.strip().strip("`")
    if not md.startswith("# "): md = f"# {slide.title}\n\n" + md
    return md

SCRIPT_TPL = ("为这页幻灯写讲课口播稿：150-300 个汉字，口语化但准确，只依据教材片段；"
              "数字一律用中文汉字。开头不要客套语。\n"
              "页标题：{title}\n幻灯内容：\n{slide}\n教材片段：{context}")

def run_script(llm, slide: OutlineSlide, slide_md: str, context: str) -> ScriptPart:
    prompt = ChatPromptTemplate.from_messages([
        ("system", STRUCT_SYS), ("human", SCRIPT_TPL)])
    msgs = prompt.format_messages(title=slide.title, slide=slide_md,
                                  context=context[:4000])
    res, err = ask_structured(llm, ScriptPart, msgs)
    if err or res is None: raise RuntimeError(f"讲稿生成失败: {err}")
    return res
```

- [ ] **Step 4: Run** → PASS
- [ ] **Step 5: 改写 tests/fakes.py**（替换 Task 4 的 `FakeStructuredLLM`；新增两个类）

```python
# tests/fakes.py
from langchain_core.messages import AIMessage

class FakeStructuredLLM:
    """按 schema 类型名返回罐头结构化结果；记录最近一次收到的消息；支持 invoke 返回纯文本"""
    def __init__(self, responses: dict):
        self.responses = responses; self.calls = 0; self.last_messages: list = []
    def with_structured_output(self, schema, **kw):
        def _fn(msgs):
            self.calls += 1; self.last_messages = list(msgs)
            if schema.__name__ not in self.responses:
                raise ValueError(f"FakeStructuredLLM 无 {schema.__name__} 罐头")
            return self.responses[schema.__name__].model_copy(deep=True)
        return _fn
    def invoke(self, msgs):
        self.calls += 1; self.last_messages = list(msgs)
        return AIMessage(content="# 罐头标题\n\n- 要点一")

class FlakyStructuredLLM(FakeStructuredLLM):
    """前 fail_times 次 with_structured_output 抛 ValueError，之后成功（测重试）"""
    def __init__(self, responses: dict, fail_times: int):
        super().__init__(responses); self.fail_times = fail_times
    def with_structured_output(self, schema, **kw):
        def _fn(msgs):
            self.calls += 1; self.last_messages = list(msgs)
            if self.calls <= self.fail_times: raise ValueError("模拟非法输出")
            return self.responses[schema.__name__].model_copy(deep=True)
        return _fn

class RawOnlyLLM:
    """with_structured_output 直接抛错，只支持 invoke 返回固定文本（测兜底路径）"""
    def __init__(self, text: str): self.text = text
    def with_structured_output(self, schema, **kw): raise RuntimeError("不支持 structured")
    def invoke(self, msgs): return AIMessage(content=self.text)
```

- [ ] **Step 6: Run** 全量 → PASS
- [ ] **Step 7: Commit** `git add -A && git commit -m "feat: LLM 工厂与三条 LCEL 链 + fake 扩展"`

---

### Task 12: 上传 pptx 转换（pipeline/pptx.py）

**Files:**
- Create: `app/pipeline/pptx.py`, `tests/test_pptx.py`

**Interfaces:**
- Consumes: `pptx_to_pngs(pptx: Path, work_dir: Path) -> list[Path]`
  - LibreOffice headless：`soffice --headless --convert-to pdf --outdir tmp pptx`（Windows 常见安装路径探测或依赖 PATH；找不到 soffice → RuntimeError 提示“请安装 LibreOffice 或改用 AI 生成 PPT”）
  - `fitz.open(pdf)` → 每页 `page.get_pixmap(dpi=150)` 存 `work_dir/slide-NN.png`（用 `page.rect` 保持 16:9）
  - 兜底 `pptx_to_text_chapters(pptx) -> str`：`python-pptx` 逐 slide 抽文本（供“重新排版模式”用）

- [ ] **Step 1: 写测试**

```python
import shutil
import pytest
from pathlib import Path
from app.pipeline import pptx as P

pytestmark = pytest.mark.skipif(shutil.which("soffice") is None and not Path(
    r"C:\Program Files\LibreOffice\program\soffice.exe").exists(), reason="no LibreOffice")

def test_pptx_to_pngs(tmp_path):
    from pptx import Presentation
    from pptx.util import Inches
    prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    for i in range(2):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        tb = s.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
        tb.text_frame.text = f"第{i+1}页"
    pptx = tmp_path / "d.pptx"; prs.save(pptx)
    pngs = P.pptx_to_pngs(pptx, tmp_path / "out")
    assert [p.name for p in pngs] == ["slide-01.png", "slide-02.png"]
    text = P.pptx_to_text_chapters(pptx)
    assert "第1页" in text
```

- [ ] **Step 2: Run** → FAIL（无 LibreOffice 时 skip）
- [ ] **Step 3: 写实现**

```python
# app/pipeline/pptx.py
import os, shutil, subprocess, tempfile
from pathlib import Path
import fitz

_SOFFICE_CANDIDATES = ["soffice", r"C:\Program Files\LibreOffice\program\soffice.exe",
                       r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]

def _soffice() -> str:
    for c in _SOFFICE_CANDIDATES:
        if c == "soffice" and shutil.which("soffice"): return "soffice"
        if c != "soffice" and Path(c).exists(): return c
    raise RuntimeError("未找到 LibreOffice（soffice）。请安装后重试，或改用 AI 生成 PPT。")

def pptx_to_pngs(pptx: Path, out_dir: Path) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp = Path(tempfile.mkdtemp(prefix="pptx2pdf-"))
    subprocess.run([_soffice(), "--headless", "--convert-to", "pdf",
                    "--outdir", str(tmp), str(pptx)], check=True,
                   capture_output=True, timeout=600)
    pdfs = list(tmp.glob("*.pdf"))
    if not pdfs: raise RuntimeError("LibreOffice 未产出 PDF")
    doc = fitz.open(pdfs[0])
    res = []
    for i, page in enumerate(doc, 1):
        dst = out_dir / f"slide-{i:02d}.png"
        page.get_pixmap(dpi=150).save(dst); res.append(dst)
    return res

def pptx_to_text_chapters(pptx: Path) -> str:
    from pptx import Presentation
    prs = Presentation(str(pptx))
    parts = [f"# 第{i}页\n" + "\n".join(
        sh.text_frame.text for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text)
        for i, slide in enumerate(prs.slides, 1)]
    return "\n\n".join(parts)
```

- [ ] **Step 4: Run** → PASS
- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: pptx→PNG（LibreOffice headless）与文本抽取兜底"`

---

### Task 13: 任务队列（pipeline/jobs.py）

**Files:**
- Create: `app/pipeline/jobs.py`, `tests/test_jobs.py`

**Interfaces:**
- Produces: 全局单例 `queue = JobQueue()`
  - `class JobStatus(str, Enum): queued running done error`
  - `class Job: id, kind: str, project_id: str, status: JobStatus, error: str|None, result: dict`
  - `class JobQueue: def __init__(self)`（内部 `asyncio.Queue` + worker task 懒启动）；`async def submit(kind, project_id, coro_factory: Callable[[], Coroutine]) -> str`（立即返回 job_id，结果/error 写回 job）；`def get(job_id) -> Job|None`；`def list() -> list[Job]`（新的在前，限 100）
  - 串行保证：单 worker 顺序执行，同一时刻最多一个重活

- [ ] **Step 1: 写测试**

```python
import asyncio
from app.pipeline.jobs import JobQueue, JobStatus

async def test_serial_execution_and_error_capture():
    q = JobQueue(); order = []
    async def slow(i):
        await asyncio.sleep(0.05); order.append(i); return {"i": i}
    async def boom(): raise RuntimeError("炸了")
    id1 = await q.submit("gen", "p1", lambda: slow(1))
    id2 = await q.submit("gen", "p1", lambda: slow(2))
    id3 = await q.submit("gen", "p1", boom)
    for _ in range(100):
        await asyncio.sleep(0.02)
        if all(q.get(i).status in (JobStatus.done, JobStatus.error) for i in (id1, id2, id3)):
            break
    assert order == [1, 2]                       # 串行：1 完成才开始 2
    assert q.get(id1).result == {"i": 1}
    assert q.get(id3).status == JobStatus.error
    assert "炸了" in q.get(id3).error
    assert q.get("nope") is None and len(q.list()) == 3
```

- [ ] **Step 2: Run** → FAIL
- [ ] **Step 3: 写实现**

```python
# app/pipeline/jobs.py
import asyncio, uuid
from collections import deque
from collections.abc import Callable, Coroutine
from enum import Enum
from typing import Any

class JobStatus(str, Enum):
    queued = "queued"; running = "running"; done = "done"; error = "error"

class Job:
    def __init__(self, kind: str, project_id: str):
        self.id = uuid.uuid4().hex[:8]; self.kind = kind; self.project_id = project_id
        self.status = JobStatus.queued; self.error: str | None = None
        self.result: dict[str, Any] = {}

class JobQueue:
    def __init__(self) -> None:
        self._q: asyncio.Queue = asyncio.Queue()
        self._jobs: dict[str, Job] = {}
        self._recent: deque[str] = deque(maxlen=100)
        self._worker: asyncio.Task | None = None
    async def submit(self, kind: str, project_id: str,
                     coro_factory: Callable[[], Coroutine]) -> str:
        job = Job(kind, project_id); self._jobs[job.id] = job
        self._recent.appendleft(job.id)
        await self._q.put((job, coro_factory))
        if self._worker is None or self._worker.done():
            self._worker = asyncio.create_task(self._run())
        return job.id
    async def _run(self) -> None:
        while True:
            job, factory = await self._q.get()
            job.status = JobStatus.running
            try:
                job.result = await factory() or {}
                job.status = JobStatus.done
            except Exception as e:              # noqa: BLE001 —— 队列吞异常写回 job
                job.status = JobStatus.error; job.error = str(e)
            finally:
                self._q.task_done()
    def get(self, job_id: str) -> Job | None: return self._jobs.get(job_id)
    def list(self) -> list[Job]: return [self._jobs[i] for i in self._recent if i in self._jobs]

queue = JobQueue()   # 全局单例；FastAPI 启动时 import 即生效
```

- [ ] **Step 4: Run** → PASS
- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: 进程内串行任务队列"`

---

### Task 14: 编排器（pipeline/orchestrator.py）

**Files:**
- Create: `app/pipeline/orchestrator.py`, `tests/test_orchestrator.py`
- Modify: `app/store.py`（追加一个路径函数：`f_parts(pid, n) -> Path` = `f_deck(pid)/"parts"/f"{sl(n)}.md"`）

**Interfaces:**（均为模块级 async 函数，读 artifacts → 写 artifacts，为将来迁 LangGraph 做纯函数纪律；store 一律用 Task 3 的模块级函数）
- 约定：材料 PDF 固定放 `store.f_raw(pid)/"material.pdf"`（Task 16 上传路由负责落盘）
- `ingest_material(pid, cfg, llm=None) -> list[Chapter]`：`make_parser(cfg, llm)` → parse 到 `store.f_parsed(pid)`；`cfg.embedding.base_url` 非空才建索引（`build_index(..., store.f_index(pid))`）；stage `ingest=done`
- `gen_outline(pid, cfg, llm) -> Outline`：拼接 parsed 下全部 md（每章前 2000 字）→ `run_outline` → outline JSON 写 `store.f_outline(pid)`；stage `outline=done`
- `confirm_slides(pid, cfg, llm) -> list[str]`：读 outline（无则先 gen_outline）→ 每页 `run_slide` 写 `store.f_parts(pid, n)` → `assemble` 写 `store.f_slides_md(pid)` → `deck.init_deck` + `render_pngs(store.f_deck(pid))` → 更新 `ProjectState.slides`（SlideMeta: index/title/png=`deck/png/slide-NN.png`）；stage `slides=done`
- `gen_page_script(pid, n, cfg, llm) -> str`：读该页 part → `retrieve(页标题)`（embedding 配置为空则 context=""）→ `run_script` → 写 `store.f_script(pid, n)`，返回文本
- `gen_page_audio(pid, n, cfg, provider=None) -> AudioResult`：读讲稿 → `synth_page(provider or EdgeTTSProvider(), text=..., voice=cfg.voice, cache_dir=store.f_cache(pid), out_mp3=store.f_audio(pid,n), out_meta=store.f_audio_meta(pid,n))` → 更新该页 SlideMeta（audio/audio_hash/audio_duration）
- `gen_page_slide(pid, n, cfg, llm) -> None`：重写第 n 页 part → 重拼 slides.md → 全 deck 重截图 → 更新该页 SlideMeta.png
- `run_project(pid, cfg, on_progress=None) -> Path`：取 ProjectState.slides 中 png+audio 齐全的页 → `build_timeline(slides, meta_of, gap=cfg.page_gap_sec, resolution=cfg.resolution)`（meta_of[n]=json.loads(store.f_audio_meta(pid,n))）→ `write_ass(tl, work_dir/"subtitles.ass")` → `render.render_video(tl, base_dir=config.project_dir(pid), work_dir=f_out(pid)/"_work", final=f_out(pid)/"final.mp4")` → `write_srt(tl, f_out(pid)/"final.srt")`；stage `render=done`

- [ ] **Step 1: 写测试**（conftest 的 autouse tmp_root 已隔离项目根；LLM/TTS/截图全用罐头，只 ffmpeg 真跑）

```python
import shutil
import fitz
import pytest
from PIL import Image
from app import config, store
from app.models import Outline, OutlineSlide, ScriptPart
from app.pipeline import orchestrator as O
from tests.fakes import FakeStructuredLLM, FakeTTSProvider

OUTLINE = Outline(slides=[OutlineSlide(title="神经元", points=["定义", "结构"]),
                          OutlineSlide(title="突触", points=["分类"])])

def _fake_llm():
    return FakeStructuredLLM({"Outline": OUTLINE,
                              "ScriptPart": ScriptPart(narration="神经元是基本单位。")})

def _material(pid):
    doc = fitz.open()
    doc.new_page().insert_text((72, 100), "神经元是大脑的基本单位。正文若干。")
    doc.save(store.f_raw(pid) / "material.pdf")

def _fake_pngs(deck_dir):
    out = deck_dir / "png"; out.mkdir(parents=True, exist_ok=True)
    res = []
    for i in (1, 2):
        p = out / f"slide-{i:02d}.png"; Image.new("RGB", (320, 180)).save(p); res.append(p)
    return res

@pytest.mark.skipif(shutil.which("ffmpeg") is None, reason="no ffmpeg")
async def test_full_pipeline_minimal(monkeypatch):
    st = store.create_project("测项目", {}); pid = st.id
    _material(pid); fake = _fake_llm()
    monkeypatch.setattr(O.deck, "render_pngs", _fake_pngs)
    monkeypatch.setattr(O, "retrieve", lambda *a, **k: [])
    monkeypatch.setattr(O, "EdgeTTSProvider", FakeTTSProvider)
    await O.ingest_material(pid, st.config, None)
    await O.gen_outline(pid, st.config, fake)
    await O.confirm_slides(pid, st.config, fake)
    await O.gen_page_script(pid, 1, st.config, fake)
    await O.gen_page_audio(pid, 1, st.config)
    final = await O.run_project(pid, st.config)
    assert final.exists() and final.stat().st_size > 0
    assert store.load_project(pid).stages["render"] == "done"
    assert (config.project_dir(pid) / "output" / "final.srt").exists()

async def test_ingest_sets_stage():
    st = store.create_project("单测", {}); _material(st.id)
    await O.ingest_material(st.id, st.config, None)
    assert store.load_project(st.id).stages["ingest"] == "done"
```

- [ ] **Step 2: Run** → FAIL
- [ ] **Step 3: 写实现**

```python
# app/pipeline/orchestrator.py
import json
from pathlib import Path
from app import config, store
from app.models import Outline, SlideMeta
from app.pipeline import deck, render
from app.pipeline.chains import run_outline, run_script, run_slide
from app.pipeline.index import build_index, make_embeddings, retrieve
from app.pipeline.ingest import make_parser
from app.pipeline.timeline import build_timeline, write_ass, write_srt
from app.pipeline.tts import EdgeTTSProvider, synth_page

def _mutate(pid, fn):
    st = store.load_project(pid); fn(st); store.save_project(st); return st

def _set_stage(pid, key, val):
    _mutate(pid, lambda st: st.stages.__setitem__(key, val))

def _upsert_slide(pid, n, **fields):
    def _fn(st):
        for s in st.slides:
            if s.index == n:
                for k, v in fields.items(): setattr(s, k, v)
                return
        st.slides.append(SlideMeta(index=n, **fields))
        st.slides.sort(key=lambda s: s.index)
    _mutate(pid, _fn)

def _outline_of(pid) -> Outline:
    return Outline.model_validate_json(store.read_text(store.f_outline(pid)))

def _chapter_text(pid) -> str:
    parts = sorted(store.f_parsed(pid).glob("*.md"))
    return "\n\n".join(p.read_text("utf-8")[:2000] for p in parts)

async def ingest_material(pid: str, cfg, llm=None) -> list:
    chapters = make_parser(cfg, llm).parse(store.f_raw(pid) / "material.pdf",
                                           store.f_parsed(pid))
    if cfg.embedding.base_url:
        build_index(chapters, make_embeddings(cfg), store.f_index(pid))
    _set_stage(pid, "ingest", "done")
    return chapters

async def gen_outline(pid: str, cfg, llm) -> Outline:
    outline = run_outline(llm, _chapter_text(pid))
    store.write_text(store.f_outline(pid), outline.model_dump_json(indent=2))
    _set_stage(pid, "outline", "done")
    return outline

async def confirm_slides(pid: str, cfg, llm) -> list[str]:
    if not store.f_outline(pid).exists():
        await gen_outline(pid, cfg, llm)
    outline = _outline_of(pid)
    deck.init_deck(store.f_deck(pid))
    parts = []
    for i, slide in enumerate(outline.slides, 1):
        part = run_slide(llm, slide, "")
        store.write_text(store.f_parts(pid, i), part)
        parts.append(part)
    store.write_text(store.f_slides_md(pid), deck.assemble(parts))
    deck.render_pngs(store.f_deck(pid))
    for i, slide in enumerate(outline.slides, 1):
        _upsert_slide(pid, i, title=slide.title, png=f"deck/png/slide-{i:02d}.png")
    _set_stage(pid, "slides", "done")
    return parts

def _has_embedding(cfg) -> bool:
    return bool(cfg.embedding.base_url)

async def gen_page_script(pid: str, n: int, cfg, llm) -> str:
    part_md = store.read_text(store.f_parts(pid, n))
    title = part_md.splitlines()[0].lstrip("# ")
    ctx = ""
    if _has_embedding(cfg):
        ctx = "\n\n".join(retrieve(title, make_embeddings(cfg), store.f_index(pid)))
    script = run_script(llm, _outline_of(pid).slides[n - 1], part_md, ctx)
    return store.write_text(store.f_script(pid, n), script.narration)

async def gen_page_audio(pid: str, n: int, cfg, provider=None):
    text = store.read_text(store.f_script(pid, n))
    res = await synth_page(provider or EdgeTTSProvider(), text=text, voice=cfg.voice,
                           cache_dir=store.f_cache(pid), out_mp3=store.f_audio(pid, n),
                           out_meta=store.f_audio_meta(pid, n))
    _upsert_slide(pid, n, audio=f"audio/{store.sl(n)}.mp3",
                  audio_hash=res.path, audio_duration=res.duration)
    return res

async def gen_page_slide(pid: str, n: int, cfg, llm) -> None:
    slide = _outline_of(pid).slides[n - 1]
    store.write_text(store.f_parts(pid, n), run_slide(llm, slide, ""))
    parts = [store.read_text(store.f_parts(pid, i))
             for i in range(1, len(_outline_of(pid).slides) + 1)
             if store.f_parts(pid, i).exists()]
    store.write_text(store.f_slides_md(pid), deck.assemble(parts))
    deck.render_pngs(store.f_deck(pid))
    _upsert_slide(pid, n, png=f"deck/png/slide-{n:02d}.png")

async def run_project(pid: str, cfg, on_progress=None) -> Path:
    st = store.load_project(pid)
    usable = [s for s in st.slides if s.png and s.audio]
    meta_of = {s.index: json.loads(store.read_text(store.f_audio_meta(pid, s.index)))
               for s in usable}
    tl = build_timeline(usable, meta_of, gap=cfg.page_gap_sec, resolution=cfg.resolution)
    work = store.f_out(pid) / "_work"; work.mkdir(parents=True, exist_ok=True)
    write_ass(tl, work / "subtitles.ass")
    final = await render.render_video(tl, base_dir=config.project_dir(pid),
                                      work_dir=work, final=store.f_out(pid) / "final.mp4",
                                      on_progress=on_progress)
    write_srt(tl, store.f_out(pid) / "final.srt")
    _set_stage(pid, "render", "done")
    return final
```

注：`render.render_video(tl, base_dir, work_dir, final)` 的 base_dir 是项目根（seg.png/seg.audio 相对其解析）；`SlideMeta.audio_hash` 存缓存文件路径仅作占位标识，实际缓存命中由 synth_page 的 cache_key 机制保证。

- [ ] **Step 4: Run** `pytest tests/test_orchestrator.py -v` → PASS（无 ffmpeg skip）
- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: 项目编排器（摄取/大纲/单页/渲染）"`

### Task 15: doctor 环境体检（__main__.py）

**Files:**
- Create: `app/__main__.py`, `tests/test_doctor.py`

**Interfaces:**
- `python -m app doctor` 输出逐项 `[ok]/[missing]`：python 版本 ≥3.11、ffmpeg、npx（marp）、soffice（LibreOffice，缺→警告不阻断）、edge-tts 连通（超时 3s 不阻断）、LLM 端点（`cfg.base_url` 可达，发一个 1-token 请求）、`MINERU_*`（仅 mineru 模式时必查）；全部必需项 ok → exit 0，否则 exit 1

- [ ] **Step 1: 写测试**

```python
from app.__main__ import checks

def test_checks_structure():
    results = checks()
    assert isinstance(results, list) and results
    for item in results:
        assert set(item) >= {"name", "ok", "required", "detail"}

def test_python_version_check():
    r = [c for c in checks() if c["name"] == "python"][0]
    assert r["ok"] is True
```

- [ ] **Step 2: Run** → FAIL
- [ ] **Step 3: 写实现**

```python
# app/__main__.py
import shutil, sys
from pathlib import Path

def checks() -> list[dict]:
    out = []
    v = sys.version_info
    out.append({"name": "python", "ok": v >= (3, 11), "required": True,
                "detail": f"{v.major}.{v.minor}.{v.micro}"})
    out.append({"name": "ffmpeg", "ok": bool(shutil.which("ffmpeg")), "required": True,
                "detail": shutil.which("ffmpeg") or "未找到"})
    out.append({"name": "npx(marp)", "ok": bool(shutil.which("npx")), "required": True,
                "detail": shutil.which("npx") or "未找到（需 Node.js）"})
    soffice = shutil.which("soffice") or Path(r"C:\Program Files\LibreOffice\program\soffice.exe").exists()
    out.append({"name": "libreoffice", "ok": bool(soffice), "required": False,
                "detail": "仅上传 pptx 时需要"})
    from app import config as C                # LLM 配置检查（不联网，只验必填非空）
    llm_cfg = C.default_project_config()["llm"]
    ok = bool(llm_cfg["base_url"] and llm_cfg["model"] and llm_cfg["api_key"])
    out.append({"name": "llm-config", "ok": ok, "required": True,
                "detail": f"{llm_cfg['model']} @ {llm_cfg['base_url']}" + ("" if ok else "（缺 api_key/model，设 LF_LLM_* 环境变量）")})
    return out

def main() -> None:
    if len(sys.argv) > 1 and sys.argv[1] == "doctor":
        bad = False
        for c in checks():
            mark = "ok     " if c["ok"] else ("MISSING" if c["required"] else "warn   ")
            print(f"[{mark}] {c['name']}: {c['detail']}")
            bad |= (not c["ok"] and c["required"])
        sys.exit(1 if bad else 0)
    print("用法: python -m app doctor")

if __name__ == "__main__": main()
```

- [ ] **Step 4: Run** `pytest tests/test_doctor.py -v && python -m app doctor` → PASS / 体检表输出符合预期
- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: doctor 环境体检"`

---

### Task 16: Web 列表页（web/main.py + templates）

**Files:**
- Create: `app/web/__init__.py`, `app/web/main.py`, `app/web/templates/index.html`

**Interfaces:**
- FastAPI app：`GET /` → Jinja2 渲染列表页（`store.list_projects()`，每行：名称/状态徽标/已用时长/最后更新 + 「新建项目」表单 POST `/projects`（multipart：title + material.pdf）→ 303 重定向）；`POST /projects/{pid}/delete` → 303
- 「状态徽标」只是 stages 汇总的只读展示（材料✓/大纲✓/音频 n/m/视频✓），无审批语义
- 启动：`uvicorn app.web.main:app --port 8765`；模板用简素内联 CSS，无 JS 构建链

- [ ] **Step 1: 写测试**（httpx TestClient）

```python
import io
import fitz
from fastapi.testclient import TestClient
from app import store
from app.web.main import app

def test_create_and_list():
    client = TestClient(app)
    doc = fitz.open(); doc.new_page().insert_text((72, 100), "hi")
    buf = io.BytesIO(); doc.save(buf)
    r = client.post("/projects", data={"title": "神经科学"},
                    files={"material": ("m.pdf", buf.getvalue(), "application/pdf")},
                    follow_redirects=False)
    assert r.status_code == 303
    page = client.get("/").text
    assert "神经科学" in page and "新建项目" in page
    assert store.list_projects()[0].title == "神经科学"

def test_delete():
    client = TestClient(app)
    st = store.create_project("x", {})
    assert client.post(f"/projects/{st.id}/delete", follow_redirects=False).status_code == 303
    assert store.list_projects() == []
```

注：项目根隔离靠 conftest 的 autouse tmp_root；路由把上传 PDF 落盘到 `store.f_raw(pid)/"material.pdf"`；`create_project` 的 cfg 用 `config.default_project_config()` 兼底。

- [ ] **Step 2: Run** → FAIL
- [ ] **Step 3: 写实现**（main.py 路由 + index.html 表单/列表， Jinja2 `{{ }}`；新建项目校验 pdf 后缀）
- [ ] **Step 4: Run** → PASS
- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: 项目列表页"`

---

### Task 17: Web 工作台（templates/workbench.html + 路由）

**Files:**
- Modify: `app/web/main.py`（增加路由）
- Create: `app/web/templates/workbench.html`

**Interfaces:**（全部 HTMX 表单/局部刷新，无前端构建链）
- `GET /projects/{pid}` → 工作台四区：
  - ①材料区：章节列表 + 摄取按钮（POST `/projects/{pid}/ingest` → 提交队列 job）
  - ②生成区：Outline 展示（只读文本框）+ 「生成大纲」`POST /outline` → 返回 outline 供确认；「确认大纲，生成幻灯」`POST /outline/confirm`（确认门默认 ON，在 UI 层）；每页一行：预览图（`/projects/{pid}/files/deck/png/slide-NN.png`）+ 讲稿 textarea（POST `/projects/{pid}/pages/{n}/script` 保存）+ 「重新生成讲稿」「重新生成幻灯」按钮 + TTS 状态
  - ③校对区：线性逐页（◀▶导航，←/→ 键盘翻页；与 ②同一页组件，翻页即切页）；保存讲稿 = 自动后台重 TTS（保存表单同时 POST script+retts）
  - ④渲染区：「生成视频」`POST /projects/{pid}/render` → job id → 轮询 `GET /jobs/{id}`（HTMX hx-get 每 2s）→ done 后给 `<video controls>` 播放 + `/projects/{pid}/files/out/final.mp4` 下载链接
- `GET /projects/{pid}/files/{path:path}`：白名单校验（解析后必须位于项目目录内，防目录穿越）
- 只有 5 个动作：摄取 / 生成大纲+确认 / 编辑保存讲稿（自动重TTS）/ 重生成单页（讲稿或幻灯）/ 生成视频；无审批状态语义，进度徽标仅只读展示

- [ ] **Step 1: 写测试**

```python
from fastapi.testclient import TestClient
from app import store
from app.web.main import app

def test_workbench_renders():
    client = TestClient(app)
    st = store.create_project("测", {})
    html = client.get(f"/projects/{st.id}").text
    assert "生成大纲" in html and "确认大纲" in html and "生成视频" in html

def test_save_script_writes_file():
    client = TestClient(app)
    st = store.create_project("测", {})
    r = client.post(f"/projects/{st.id}/pages/1/script",
                    data={"text": "新讲稿。"}, follow_redirects=False)
    assert r.status_code == 303
    assert store.read_text(store.f_script(st.id, 1)) == "新讲稿。"

def test_files_route_blocks_traversal():
    client = TestClient(app)
    st = store.create_project("测", {})
    r = client.get(f"/projects/{st.id}/files/..%2F..%2Fetc%2Fpasswd")
    assert r.status_code in (400, 404)
```

注：保存讲稿路由内部会用 `jobs.queue.submit` 异步重跑该页 TTS（工作台真实流程）；单测只验证写盘与 303，不等待 job 完成。

- [ ] **Step 2: Run** → FAIL
- [ ] **Step 3: 写实现**（路由 ~8 个 + workbench.html；HTMX 片段路由：`GET /projects/{pid}/pages/{n}/panel` 返回单页 panel HTML，翻页即换片段）
- [ ] **Step 4: Run** 全量 → PASS
- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: 工作台四区（HTMX）与文件服务白名单"`

---

### Task 18: README、冒烟脚本与收尾

**Files:**
- Create: `README.md`, `scripts/smoke.py`
- Modify: `pyproject.toml`（补 `[project.scripts] lessonforge = "app.__main__:main"`）

**Interfaces:**
- `scripts/smoke.py`：端到端冒烟（真实 LLM/edge-tts/ffmpeg/marp，需 `--yes` 跳过提示）：建项目（用 `examples/mini.pdf` 3 页）→ ingest → outline（打印确认 y/n）→ confirm → 全页 script+audio → render → 打印 final 路径与时长
- README：安装（uv sync / pip install -e .）、外部依赖表（ffmpeg/Node+marp/LibreOffice/MinerU 可选）、`python -m app doctor`、快速开始（4 步）、目录结构说明、常见错误（ffmpeg 缺失、soffice 缺失、LLM 端点 401、marp 首次跑慢）

- [ ] **Step 1: 写 examples/mini.pdf 生成器**（脚本内嵌 fitz 生成 3 页中文文本页）
- [ ] **Step 2: Run** `python scripts/smoke.py --yes` 手动验证一次（本任务唯一非自动化步骤，通过标准：产出可播放 final.mp4 且字幕同步）
- [ ] **Step 3: 写 README.md**（结构见上；所有命令复制即用）
- [ ] **Step 4: 全量回归**：`pytest -q` 全绿；`git log --oneline` 任务提交历史完整
- [ ] **Step 5: Commit** `git add -A && git commit -m "docs: README 与端到端冒烟脚本"`

---

## 附：Spec 覆盖对照

| Spec 章节 | 落点任务 |
|---|---|
| §4 数据模型/目录布局 | Task 1, 2, 3 |
| §5 管线五阶段 | Task 8, 9（摄取）、11（大纲/幻灯/讲稿）、4（音频）、6+14（渲染）、5（时间轴） |
| §6 字幕机制 | Task 5, 6 |
| §7 校对界面 | Task 16, 17 |
| §8 错误处理 | Task 8（LLM 兜底→单章）、11（ask_structured 重试）、12（pptx 兜底）、13（job error）、15（doctor） |
| §9 测试 | Task 1（conftest/fakes）+ 各任务 TDD + Task 18（smoke） |
| §10 升级路径 | 不在本计划内（仅保证 Timeline/渲染层预留 bgm/overlay 字段，Task 2 已含） |
| §11 LangGraph 迁移触发条件 | Task 14 纯函数式编排即为此准备 |
| §12 环境依赖 | Task 15 doctor、Task 18 README |

## 附：执行顺序与并行性

严格按 Task 1→18 顺序执行（TDD 依赖链）。全串行，无并行任务；单人单机一次只跟一个任务。

每任务完成标准：新测试先 FAIL 后 PASS + 全量 pytest 绿 + 单独 commit。若某任务卡住超过半小时，在 plan 对应任务下追加「阻塞记录」小节，不静默跳过。
