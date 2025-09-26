import os
import re
import json
from run_complete_analysis import log_message, _safe_console_print
from app.services.matching_service import MatchingService

# Basic smoke tests for safe logging and matching service error structure.

def test_safe_console_print_does_not_raise():
    # This includes characters that previously caused UnicodeEncodeError on cp1252 consoles
    dangerous = "Test symbols ✓ 🎉 ✅ — end"
    _safe_console_print(dangerous)


def test_log_message_writes_utf8(tmp_path, monkeypatch):
    log_file = tmp_path / 'test_log.txt'
    # Monkeypatch Output directory path by changing CWD temporarily
    cwd = os.getcwd()
    try:
        os.chdir(tmp_path)
        log_message("Unicode ✓ 🎉 line", log_file=str(log_file))
        data = log_file.read_text(encoding='utf-8')
        assert 'Unicode' in data
    finally:
        os.chdir(cwd)


def test_matching_service_failure_structure(monkeypatch):
    # Force subprocess to fail by pointing to a non-existent script via monkeypatch
    ms = MatchingService()

    # Monkeypatch _run_complete_analysis to simulate failure return structure
    def fake_run_complete_analysis(self, use_flexible=False):
        return { 'success': False, 'message': 'Analysis failed', 'stderr': 'stderr text', 'stdout': 'stdout text' }

    monkeypatch.setattr(MatchingService, '_run_complete_analysis', fake_run_complete_analysis)

    result = ms.run_matching()
    assert result['success'] is False
    assert 'stderr' in result or 'message' in result
