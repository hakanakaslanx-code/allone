import importlib
import sys
import types
from pathlib import Path

import pytest


def _ensure_repo_on_path() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    repo_str = str(repo_root)
    if repo_str not in sys.path:
        sys.path.insert(0, repo_str)


def _ensure_requests_stub() -> None:
    if "requests" in sys.modules:
        return
    stub = types.ModuleType("requests")

    class Timeout(Exception):
        pass

    class RequestException(Exception):
        pass

    stub.Timeout = Timeout
    stub.RequestException = RequestException

    def fake_get(*_args, **_kwargs):  # pragma: no cover - defensive stub
        raise Timeout()

    stub.get = fake_get
    sys.modules["requests"] = stub


@pytest.fixture()
def updater_module():
    _ensure_repo_on_path()
    _ensure_requests_stub()
    return importlib.reload(importlib.import_module("allone.updater"))


def test_parse_version_handles_prefixed_tags(updater_module):
    parse = updater_module._parse_version
    assert parse("v5.1.3") == ((1, 5), (1, 1), (1, 3))
    assert parse("release-2024.09.1-beta") == (
        (1, 2024),
        (1, 9),
        (1, 1),
        (0, "beta"),
    )
    assert parse("beta") == ()


def test_version_comparison_with_typeerror_fallback(monkeypatch, updater_module):
    def legacy_parse(value: str):
        parts = []
        for chunk in value.split('.'):
            if chunk.isdigit():
                parts.append(int(chunk))
            else:
                parts.append(chunk)
        return tuple(parts)

    monkeypatch.setattr(updater_module, "_parse_version", legacy_parse)

    is_not_newer = updater_module._version_is_not_newer

    assert is_not_newer("v5.1.3", "5.1.3")
    assert not is_not_newer("v5.2.0", "5.1.3")
    assert is_not_newer("v5.1.3-beta", "5.1.3")
