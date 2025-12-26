import importlib
import sys
import types
from pathlib import Path


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


def test_update_cli_resolves_release_without_current_json(monkeypatch, tmp_path):
    _ensure_repo_on_path()
    _ensure_requests_stub()
    update_cli = importlib.reload(importlib.import_module("allone.update_cli"))

    release_dir = tmp_path / "releases" / "9.9.9"
    release_dir.mkdir(parents=True)
    release_exe = release_dir / "AllOne Tools.exe"
    release_exe.write_text("stub exe")

    monkeypatch.setattr(update_cli, "_remote_version_string", lambda _timeout: "0.0.0")
    monkeypatch.setattr(
        update_cli, "_version_is_not_newer", lambda _remote, _current: True
    )

    exit_code = update_cli.main(["--install-root", str(tmp_path)])

    assert exit_code == 0
