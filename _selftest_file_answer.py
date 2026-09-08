import os
import tempfile
from types import SimpleNamespace

import file_answer
from file_answer import build_question_prompt, read_small_text_file


def test_read_small_text_file_returns_content_and_identifier():
    with tempfile.TemporaryDirectory() as temp_dir:
        path = os.path.join(temp_dir, "메모.txt")
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("회의 시작: 오후 3시")

        result = read_small_text_file(path)

        assert result["text"] == "회의 시작: 오후 3시"
        assert result["file_name"] == "메모.txt"


def test_read_small_text_file_rejects_unsupported_extension():
    with tempfile.TemporaryDirectory() as temp_dir:
        path = os.path.join(temp_dir, "메모.csv")
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("회의 시작: 오후 3시")

        try:
            read_small_text_file(path)
        except ValueError as error:
            assert "지원하지 않는 파일 형식" in str(error)
        else:
            raise AssertionError("지원하지 않는 확장자를 허용함")


def test_read_small_text_file_rejects_missing_file():
    try:
        read_small_text_file("존재하지 않는 파일.txt")
    except FileNotFoundError:
        pass
    else:
        raise AssertionError("없는 파일을 허용함")


def test_read_small_text_file_rejects_oversized_file():
    with tempfile.TemporaryDirectory() as temp_dir:
        path = os.path.join(temp_dir, "큰파일.txt")
        with open(path, "wb") as handle:
            handle.write(b"x" * (file_answer.MAX_FILE_BYTES + 1))

        try:
            read_small_text_file(path)
        except ValueError as error:
            assert "1 MiB" in str(error)
        else:
            raise AssertionError("크기 초과 파일을 허용함")


def test_build_question_prompt_contains_file_text_and_question():
    prompt = build_question_prompt("회의 시작: 오후 3시", "회의 시작 시간이 언제야?")

    assert "회의 시작: 오후 3시" in prompt
    assert "회의 시작 시간이 언제야?" in prompt


def test_answer_file_question_passes_file_context_to_model():
    with tempfile.TemporaryDirectory() as temp_dir:
        path = os.path.join(temp_dir, "메모.md")
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("회의 시작: 오후 3시")

        calls = []

        def fake_chat(**kwargs):
            calls.append(kwargs)
            return {"message": {"content": "오후 3시입니다."}}

        original_chat = file_answer.ollama.chat
        file_answer.ollama.chat = fake_chat
        try:
            answer = file_answer.answer_file_question(path, "회의 시작 시간이 언제야?")
        finally:
            file_answer.ollama.chat = original_chat

        assert answer == "오후 3시입니다."
        assert "회의 시작: 오후 3시" in calls[0]["messages"][0]["content"]


def test_answer_file_question_rejects_empty_model_response():
    with tempfile.TemporaryDirectory() as temp_dir:
        path = os.path.join(temp_dir, "메모.txt")
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("내용")

        original_chat = file_answer.ollama.chat
        original_requests = getattr(file_answer, "requests", None)
        file_answer.ollama.chat = lambda **kwargs: {"message": {"content": ""}}
        file_answer.requests = SimpleNamespace(
            post=lambda *args, **kwargs: (_ for _ in ()).throw(ConnectionError("vllm unavailable"))
        )
        try:
            try:
                file_answer.answer_file_question(path, "무슨 내용이야?")
            except RuntimeError as error:
                assert "Ollama" in str(error)
                assert "vLLM" in str(error)
            else:
                raise AssertionError("두 provider 실패를 성공으로 처리함")
        finally:
            file_answer.ollama.chat = original_chat
            if original_requests is None:
                del file_answer.requests
            else:
                file_answer.requests = original_requests


def test_answer_file_question_falls_back_to_vllm_when_ollama_fails():
    with tempfile.TemporaryDirectory() as temp_dir:
        path = os.path.join(temp_dir, "메모.txt")
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("회의 시작: 오후 3시")

        def failing_ollama(**kwargs):
            raise ConnectionError("ollama unavailable")

        class FakeResponse:
            def raise_for_status(self):
                return None

            def json(self):
                return {"choices": [{"message": {"content": "오후 3시입니다."}}]}

        original_chat = file_answer.ollama.chat
        original_requests = getattr(file_answer, "requests", None)
        calls = []
        file_answer.ollama.chat = failing_ollama
        file_answer.requests = SimpleNamespace(
            post=lambda *args, **kwargs: (calls.append((args, kwargs)) or FakeResponse())
        )
        try:
            answer = file_answer.answer_file_question(path, "회의 시작 시간이 언제야?")
        finally:
            file_answer.ollama.chat = original_chat
            if original_requests is None:
                del file_answer.requests
            else:
                file_answer.requests = original_requests

        assert answer == "오후 3시입니다."
        assert calls[0][0][0].endswith("/v1/chat/completions")
        assert calls[0][1]["json"]["model"] == file_answer.VLLM_MODEL
        assert calls[0][1]["json"]["chat_template_kwargs"] == {"enable_thinking": False}


def test_answer_file_question_reports_both_provider_failures():
    with tempfile.TemporaryDirectory() as temp_dir:
        path = os.path.join(temp_dir, "메모.txt")
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("내용")

        original_chat = file_answer.ollama.chat
        original_requests = getattr(file_answer, "requests", None)
        file_answer.ollama.chat = lambda **kwargs: (_ for _ in ()).throw(ConnectionError("ollama unavailable"))
        file_answer.requests = SimpleNamespace(
            post=lambda *args, **kwargs: (_ for _ in ()).throw(ConnectionError("vllm unavailable"))
        )
        try:
            try:
                file_answer.answer_file_question(path, "무슨 내용이야?")
            except RuntimeError as error:
                assert "Ollama" in str(error)
                assert "vLLM" in str(error)
            else:
                raise AssertionError("두 provider 실패를 성공으로 처리함")
        finally:
            file_answer.ollama.chat = original_chat
            if original_requests is None:
                del file_answer.requests
            else:
                file_answer.requests = original_requests


if __name__ == "__main__":
    test_read_small_text_file_returns_content_and_identifier()
    test_read_small_text_file_rejects_unsupported_extension()
    test_read_small_text_file_rejects_missing_file()
    test_read_small_text_file_rejects_oversized_file()
    test_build_question_prompt_contains_file_text_and_question()
    test_answer_file_question_passes_file_context_to_model()
    test_answer_file_question_rejects_empty_model_response()
    test_answer_file_question_falls_back_to_vllm_when_ollama_fails()
    test_answer_file_question_reports_both_provider_failures()
    print("file_answer self-test 통과")