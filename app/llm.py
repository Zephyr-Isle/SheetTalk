"""OpenAI 兼容客户端。支持流式输出(on_delta 回调)与 function calling。
兼容 DeepSeek/智谱/月之暗面/通义/OpenAI/Ollama 等。
"""
import json

import requests


class LLMError(Exception):
    pass


def _headers(api_key):
    h = {"Content-Type": "application/json"}
    if api_key:
        h["Authorization"] = "Bearer " + api_key
    return h


def _accumulate_stream(resp, on_delta):
    """解析 SSE 流,边收边回调内容增量,最终组装出完整 message(含 tool_calls)。"""
    content_parts = []
    tool_calls = {}  # index -> {"id","name","args"}
    for raw in resp.iter_lines():
        if not raw:
            continue
        if isinstance(raw, bytes):
            # SSE 响应头常不带 charset,requests 会误按 latin-1 解码导致中文乱码
            line = raw.decode("utf-8", "replace").strip()
        else:
            line = raw.strip()
        if not line.startswith("data:"):
            continue
        data = line[5:].strip()
        if data == "[DONE]":
            break
        try:
            obj = json.loads(data)
        except Exception:
            continue
        choices = obj.get("choices") or []
        if not choices:
            continue
        delta = choices[0].get("delta") or {}
        if delta.get("content"):
            content_parts.append(delta["content"])
            if on_delta:
                try:
                    on_delta(delta["content"])
                except Exception:
                    pass
        for tc in delta.get("tool_calls") or []:
            i = tc.get("index") or 0
            acc = tool_calls.setdefault(i, {"id": "", "name": "", "args": ""})
            if tc.get("id"):
                acc["id"] = tc["id"]
            fn = tc.get("function") or {}
            if fn.get("name"):
                acc["name"] = fn["name"]
            if fn.get("arguments"):
                acc["args"] += fn["arguments"]
    message = {"content": "".join(content_parts) or "", "tool_calls": None}
    if tool_calls:
        message["tool_calls"] = [
            {"id": v["id"] or f"call_{k}", "type": "function",
             "function": {"name": v["name"], "arguments": v["args"] or "{}"}}
            for k, v in sorted(tool_calls.items())]
    return message


def chat_completions(base_url, api_key, model, messages, tools=None, temperature=0.3,
                     max_tokens=None, timeout=300, on_delta=None):
    url = base_url.rstrip("/") + "/chat/completions"
    payload = {"model": model, "messages": messages, "temperature": temperature}
    if max_tokens:
        payload["max_tokens"] = max_tokens
    if tools:
        payload["tools"] = tools
        payload["tool_choice"] = "auto"

    if on_delta is not None:
        # 流式优先;部分网关不支持流式时回退非流式
        stream_payload = dict(payload, stream=True)
        try:
            resp = requests.post(url, headers=_headers(api_key), json=stream_payload,
                                 stream=True, timeout=timeout)
        except requests.RequestException as e:
            raise LLMError(f"无法连接模型服务 {url}:{e}")
        if resp.status_code == 200:
            try:
                return _accumulate_stream(resp, on_delta)
            except requests.RequestException as e:
                raise LLMError(f"模型流式连接中断:{e}")
        if resp.status_code != 400:
            raise LLMError(f"模型服务返回 HTTP {resp.status_code}:{resp.text[:400]}")
        try:
            err_text = resp.text[:200]
        except Exception:
            err_text = ""
        if "stream" not in err_text.lower() and "sse" not in err_text.lower():
            raise LLMError(f"模型服务返回 HTTP 400:{err_text}")
        resp.close()

    # 非流式
    try:
        resp = requests.post(url, headers=_headers(api_key), json=payload, timeout=timeout)
    except requests.RequestException as e:
        raise LLMError(f"无法连接模型服务 {url}:{e}")
    if resp.status_code != 200:
        raise LLMError(f"模型服务返回 HTTP {resp.status_code}:{resp.text[:400]}")
    try:
        data = resp.json()
        msg = data["choices"][0]["message"]
        msg.setdefault("content", "")
        msg.setdefault("tool_calls", None)
    except Exception:
        raise LLMError("模型返回格式异常:" + resp.text[:400])
    if on_delta and msg.get("content"):
        on_delta(msg["content"])
    return msg


def is_tool_unsupported_error(err_text):
    """判断错误是否像「模型不支持 function calling」。

    只认工具相关关键词,不认裸 400/422:否则模型名填错之类的普通 HTTP 400
    也会被误判,给用户「请换模型」的无关提示。
    """
    t = str(err_text).lower()
    return any(k in t for k in ("tool", "function", "function calling"))
