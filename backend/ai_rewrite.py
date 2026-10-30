# -*- coding: utf-8 -*-
"""调用 OpenAI 兼容接口，把 Git 提交改写成周报任务描述。"""

import re

import requests

DEFAULT_TIMEOUT = 60


def _endpoint(base_url):
    base = (base_url or "").strip().rstrip("/")
    if not base:
        raise RuntimeError("未配置 AI Base URL")
    if base.endswith(("/chat/completions", "/responses", "/messages")):
        return base
    return base + "/chat/completions"


def _headers(options):
    key = (options.get("ai_api_key") or "").strip()
    if not key:
        raise RuntimeError("未配置 AI API Key")
    return {
        "Authorization": "Bearer " + key,
        "Content-Type": "application/json",
    }


def _extract_text(data):
    if not isinstance(data, dict):
        return ""
    choices = data.get("choices")
    if choices:
        choice = choices[0] or {}
        message = choice.get("message") or {}
        content = message.get("content")
        if isinstance(content, str) and content.strip():
            return content
        if isinstance(content, list):
            parts = [
                part.get("text", "")
                for part in content
                if isinstance(part, dict) and part.get("text")
            ]
            if parts:
                return "".join(parts)
        if isinstance(choice.get("text"), str):
            return choice["text"]
    content = data.get("content")
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = [
            part.get("text", "")
            for part in content
            if isinstance(part, dict) and part.get("text")
        ]
        if parts:
            return "".join(parts)
    for key in ("output_text", "text", "response"):
        if isinstance(data.get(key), str):
            return data[key]
    output = data.get("output")
    if isinstance(output, list):
        parts = []
        for item in output:
            if not isinstance(item, dict):
                continue
            for content_item in item.get("content", []) or []:
                if isinstance(content_item, dict) and content_item.get("text"):
                    parts.append(content_item["text"])
        if parts:
            return "".join(parts)
    return ""


def _clean(text):
    text = str(text or "").strip()
    text = re.sub(r"^```[a-zA-Z0-9_-]*\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    text = text.strip().strip('"').strip("'").strip()
    text = re.sub(r"^(任务描述|描述|周报描述|完成事项)[:：]\s*", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _post(messages, options):
    url = _endpoint(options.get("ai_base_url"))
    model = (options.get("ai_model") or "").strip()
    if not model:
        raise RuntimeError("未配置 AI 模型")
    payload = {
        "model": model,
        "messages": messages,
        "temperature": float(options.get("ai_temperature") or 0.3),
        "max_tokens": 500,
    }
    timeout = int(options.get("ai_timeout") or DEFAULT_TIMEOUT)
    try:
        response = requests.post(
            url, headers=_headers(options), json=payload, timeout=timeout
        )
    except requests.RequestException as error:
        raise RuntimeError("AI 请求失败: {}".format(error))
    if response.status_code >= 400:
        raise RuntimeError(
            "AI 接口返回 {}: {}".format(response.status_code, response.text[:400])
        )
    try:
        data = response.json()
    except ValueError:
        raise RuntimeError("AI 接口返回的不是 JSON")
    text = _extract_text(data)
    if not text:
        raise RuntimeError("AI 接口没有返回文本")
    return _clean(text)


def _subjects(row):
    subjects = [str(item).strip() for item in (row.get("subjects") or [])]
    subjects = [item for item in subjects if item]
    if subjects:
        return subjects
    return [
        item.strip()
        for item in str(row.get("description") or "").split("；")
        if item.strip()
    ]


def _build_messages(row, options):
    prompt = (options.get("ai_prompt") or "").strip()
    if not prompt:
        prompt = (
            "你是周报助手。根据同一个 Ticket 下的 Git 提交记录，写一条简洁、"
            "专业的中文任务描述，用于周报的「任务描述」列。只输出描述正文。"
        )
    lines = [
        "Ticket: " + str(row.get("ticket") or "未分类"),
        "仓库: " + str(row.get("repos") or ""),
        "分支: " + str(row.get("branches") or ""),
        "日期: " + str(row.get("date_range") or ""),
        "提交记录:",
    ]
    for subject in _subjects(row)[:80]:
        lines.append("- " + subject)
    return [
        {"role": "system", "content": prompt},
        {"role": "user", "content": "\n".join(lines)},
    ]


def rewrite_row(row, options):
    text = _post(_build_messages(row, options), options)
    max_chars = int(options.get("max_description_chars") or 1000)
    if max_chars and len(text) > max_chars:
        text = text[: max_chars - 1] + "…"
    return text


def rewrite_rows(rows, options, warnings=None):
    warnings = warnings if warnings is not None else []
    if not options.get("ai_enabled"):
        return rows
    for row in rows:
        try:
            text = rewrite_row(row, options)
            if text:
                row["description"] = text
                row["ai_generated"] = True
        except Exception as error:  # noqa: BLE001
            warnings.append(
                "AI 改写失败（{}）：{}".format(row.get("ticket") or "未分类", error)
            )
    return rows


def test_connection(options):
    reply = _post(
        [
            {"role": "system", "content": "你是连通性测试助手。"},
            {"role": "user", "content": "请只回复四个字：连接成功"},
        ],
        options,
    )
    return {
        "ok": True,
        "reply": reply,
        "model": options.get("ai_model"),
        "endpoint": _endpoint(options.get("ai_base_url")),
    }
