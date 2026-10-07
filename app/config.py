"""配置读写。开发环境存放在项目根目录 config.json,打包后存放在 %APPDATA%/ExcelAI。"""
import json
import os
import sys
import threading

from app import runtime

PRESETS = {
    "deepseek": {"label": "DeepSeek(推荐)", "base_url": "https://api.deepseek.com", "model": "deepseek-chat"},
    "zhipu": {"label": "智谱 GLM", "base_url": "https://open.bigmodel.cn/api/paas/v4", "model": "glm-4.6"},
    "moonshot": {"label": "月之暗面 Kimi", "base_url": "https://api.moonshot.cn/v1", "model": "kimi-k2-0905-preview"},
    "dashscope": {"label": "阿里通义千问", "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1", "model": "qwen-plus"},
    "openai": {"label": "OpenAI", "base_url": "https://api.openai.com/v1", "model": "gpt-4o"},
    "ollama": {"label": "本地 Ollama(无需密钥)", "base_url": "http://127.0.0.1:11434/v1", "model": "qwen2.5:7b"},
    "custom": {"label": "自定义(OpenAI 兼容接口)", "base_url": "", "model": ""},
}

DEFAULTS = {
    "provider": "deepseek",
    "base_url": PRESETS["deepseek"]["base_url"],
    "model": PRESETS["deepseek"]["model"],
    "api_key": "",
    "theme": "light",        # 旧键,保留兼容;界面主题已改用下面的 theme_mode
    "theme_mode": "auto",    # 界面主题:auto=跟随系统(默认)/ light / dark
    "samples": None,         # 输入框上方示例按钮;None=内置默认,[]=不显示(用户删光)
    "max_steps": 30,
    # 计划模式:先让模型列执行计划再动手(对齐 Copilot 的 Plan 模式),默认关闭
    "plan_mode": False,
    # 工具 RAG:每轮只下发「常驻核心 + 检索命中 + 元工具」;false=全量 74 个(费 token)
    "tool_rag": True,
    # 危险工具(删表/清空/关簿等)执行前向用户发起二次确认
    "confirm_dangerous": True,
}

_lock = threading.Lock()


def _config_path():
    if runtime.is_frozen():
        base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "ExcelAI")
        os.makedirs(base, exist_ok=True)
        return os.path.join(base, "config.json")
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config.json")


def data_dir():
    """配置文件所在目录:开发态=项目根,打包态=%APPDATA%/ExcelAI。

    供 config.json 以外的运行痕迹文件使用(如「用户主动退出」标记),别存大数据。
    """
    return os.path.dirname(_config_path())


def _read_raw():
    try:
        with open(_config_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _merge(data):
    s = dict(DEFAULTS)
    s.update({k: v for k, v in (data or {}).items() if k in DEFAULTS})
    return s


def load():
    with _lock:
        return _merge(_read_raw())


def save(patch):
    with _lock:
        data = _read_raw()
        for k, v in (patch or {}).items():
            if k in DEFAULTS:
                data[k] = v
        try:
            with open(_config_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        return _merge(data)
