"""Translate provider APIs into the agent's small, explicit message contract."""
from __future__ import annotations

import json
import uuid
from typing import Any

from .agent import ModelReply, ToolRequest, ToolSpec


def arguments_json(value: dict | str) -> str:
    return value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)


class OpenAIProvider:
    """Stateless Responses requests; retain reasoning items for continuation."""
    def __init__(self, client, model: str, max_output_tokens: int = 4096):
        self.client, self.model = client, model
        self.max_output_tokens = max_output_tokens

    @staticmethod
    def input_items(messages: list[dict]) -> list[dict]:
        items = []
        for message in messages:
            if message["role"] == "tool":
                items.append({"type": "function_call_output", "call_id": message["tool_call_id"], "output": message["content"]})
            elif message["role"] == "assistant":
                if message.get("provider_items"):
                    items.extend(message["provider_items"])
                else:
                    if message.get("content"):
                        items.append({"role": "assistant", "content": message["content"]})
                    for call in message.get("tool_requests", []):
                        items.append({"type": "function_call", "call_id": call["id"], "name": call["name"], "arguments": arguments_json(call["arguments"])})
            else:
                items.append({"role": message["role"], "content": message["content"]})
        return items

    async def complete(self, messages: list[dict], tools: list[ToolSpec]) -> ModelReply:
        response = await self.client.responses.create(
            model=self.model, input=self.input_items(messages), store=False,
            include=["reasoning.encrypted_content"], max_output_tokens=self.max_output_tokens,
            tools=[{"type": "function", "name": t.name, "description": t.description,
                    "parameters": t.parameters, "strict": False} for t in tools],
        )
        if response.status != "completed":
            raise ValueError("Incomplete Responses output; no tool will be executed")
        calls, text, raw = [], [], []
        for item in response.output:
            data = item.model_dump(exclude_none=True)
            raw.append(data)
            if item.type == "function_call":
                calls.append(ToolRequest(item.call_id, item.name, item.arguments))
            elif item.type == "message":
                for part in item.content:
                    if part.type == "output_text":
                        text.append(part.text)
            elif item.type != "reasoning":
                raise ValueError("Unsupported Responses output type")
        usage = response.usage
        return ModelReply("\n".join(text), calls,
                          {"input_tokens": usage.input_tokens, "output_tokens": usage.output_tokens} if usage else {}, raw)


class VLLMProvider:
    """OpenAI-compatible Chat Completions API (self-hosted vLLM), distinct from
    OpenAIProvider's Responses API above - a real vLLM/Colab endpoint answers
    /v1/chat/completions, not /v1/responses, so OpenAIProvider cannot be
    pointed at it just by changing base_url. chat_template_kwargs suppresses
    a <think> preamble some Qwen builds emit by default (verified against a
    live vLLM endpoint in this project's chat_assistant integration)."""
    def __init__(self, client, model: str, max_output_tokens: int = 4096):
        self.client, self.model = client, model
        self.max_output_tokens = max_output_tokens

    @staticmethod
    def input_messages(messages: list[dict]) -> list[dict]:
        result = []
        for message in messages:
            if message["role"] == "tool":
                result.append({"role": "tool", "tool_call_id": message["tool_call_id"], "content": message["content"]})
            elif message["role"] == "assistant":
                entry: dict[str, Any] = {"role": "assistant", "content": message.get("content") or None}
                if message.get("tool_requests"):
                    entry["tool_calls"] = [
                        {"id": call["id"], "type": "function",
                         "function": {"name": call["name"], "arguments": arguments_json(call["arguments"])}}
                        for call in message["tool_requests"]
                    ]
                result.append(entry)
            else:
                result.append({"role": message["role"], "content": message["content"]})
        return result

    async def complete(self, messages: list[dict], tools: list[ToolSpec]) -> ModelReply:
        response = await self.client.post("/chat/completions", json={
            "model": self.model, "stream": False,
            "messages": self.input_messages(messages),
            "max_tokens": self.max_output_tokens,
            "chat_template_kwargs": {"enable_thinking": False},
            "tools": [{"type": "function", "function": {"name": t.name, "description": t.description, "parameters": t.parameters}} for t in tools],
        }, headers={"ngrok-skip-browser-warning": "1"})
        response.raise_for_status()
        data = response.json()
        choice = data["choices"][0]
        if choice.get("finish_reason") not in {"stop", "tool_calls", None}:
            raise ValueError(f"Incomplete vLLM output ({choice.get('finish_reason')}); no tool will be executed")
        message = choice["message"]
        requests = [
            ToolRequest(call.get("id") or "local-" + uuid.uuid4().hex,
                        call["function"]["name"], call["function"]["arguments"])
            for call in message.get("tool_calls") or []
        ]
        usage = data.get("usage") or {}
        return ModelReply(message.get("content") or "", requests,
                          {key: usage[source] for key, source in (("input_tokens", "prompt_tokens"), ("output_tokens", "completion_tokens")) if source in usage}, [])


class OllamaProvider:
    """Native /api/chat, not an assumption of full OpenAI compatibility."""
    def __init__(self, client, model: str, max_output_tokens: int = 4096):
        self.client, self.model = client, model
        self.max_output_tokens = max_output_tokens

    @staticmethod
    def input_messages(messages: list[dict]) -> list[dict]:
        result = []
        for message in messages:
            entry: dict[str, Any] = {"role": message["role"], "content": message.get("content", "")}
            if message["role"] == "tool":
                entry["tool_name"] = message["name"]
            if message.get("tool_requests"):
                calls = []
                for call in message["tool_requests"]:
                    args = call["arguments"]
                    if isinstance(args, str):
                        args = json.loads(args)
                    calls.append({"function": {"name": call["name"], "arguments": args}})
                entry["tool_calls"] = calls
            for opaque in message.get("provider_items", []):
                if "thinking" in opaque:
                    entry["thinking"] = opaque["thinking"]
            result.append(entry)
        return result

    async def complete(self, messages: list[dict], tools: list[ToolSpec]) -> ModelReply:
        response = await self.client.post("/api/chat", json={
            "model": self.model, "stream": False, "messages": self.input_messages(messages),
            "options": {"num_predict": self.max_output_tokens},
            "tools": [{"type": "function", "function": {"name": t.name, "description": t.description, "parameters": t.parameters}} for t in tools],
        })
        response.raise_for_status()
        data = response.json()
        if not data.get("done") or data.get("done_reason") not in {None, "stop"}:
            raise ValueError("Incomplete Ollama output; no tool will be executed")
        message = data["message"]
        requests = []
        for call in message.get("tool_calls", []):
            function = call["function"]
            # Native Ollama may not issue a call ID; app IDs pair its own results.
            requests.append(ToolRequest(call.get("id") or "local-" + uuid.uuid4().hex,
                                        function["name"], function["arguments"]))
        opaque = [{"thinking": message["thinking"]}] if message.get("thinking") else []
        return ModelReply(message.get("content", ""), requests,
                          {key: data[source] for key, source in (("input_tokens", "prompt_eval_count"), ("output_tokens", "eval_count")) if source in data}, opaque)
