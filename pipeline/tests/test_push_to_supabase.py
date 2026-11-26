"""
Tests for push_to_supabase.py's self-healing trim/retry logic, added
26-Nov-2026 after a real GitHub Actions failure: the `dashboard_payload`
table had never been trimmed (every daily run just added a row), and grew
large enough that the INSERT itself started hitting a Postgres statement
timeout (HTTP 500, code 57014). These tests mock urllib.request.urlopen
entirely -- no real network calls -- and assert on the SEQUENCE of
requests this module makes (method + URL), since that sequence is the
actual behavior being fixed: insert, then trim down to one row on the
normal path; and on a 57014 timeout, trim-keep-newest then retry once
(never leaving the table with zero rows, even if the retry also fails).
"""
import io
import json
import os
import sys
import unittest
import urllib.error
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import push_to_supabase as pts


def _resp(status, body_obj):
    """A fake context-manager response, like what urlopen(...) as resp gives."""
    m = MagicMock()
    m.__enter__.return_value = m
    m.__exit__.return_value = False
    m.status = status
    m.read.return_value = json.dumps(body_obj).encode("utf-8") if body_obj is not None else b""
    return m


def _http_error(code, body_obj):
    fp = io.BytesIO(json.dumps(body_obj).encode("utf-8"))
    return urllib.error.HTTPError("https://example.supabase.co/rest/v1/dashboard_payload",
                                   code, "err", {}, fp)


TIMEOUT_BODY = {"code": "57014", "message": "canceling statement due to statement timeout"}


class _PushTestBase(unittest.TestCase):
    """push() reads SUPABASE_URL/SUPABASE_SERVICE_ROLE_KEY from the
    environment before making any request -- these tests only care about
    the request sequence the function makes, so a fixed fake pair is
    enough, scoped to each test via patch.dict (never left in the real
    environment afterward)."""

    def setUp(self):
        patcher = patch.dict(os.environ, {
            "SUPABASE_URL": "https://x.supabase.co",
            "SUPABASE_SERVICE_ROLE_KEY": "test-service-role-key",
        })
        patcher.start()
        self.addCleanup(patcher.stop)


class TestPushNormalPath(_PushTestBase):
    """The common day-to-day case: insert succeeds, then this new trimming
    step keeps the table at exactly 1 row (the fix for the underlying
    unbounded-growth problem, not just the acute timeout symptom)."""

    def setUp(self):
        super().setUp()
        self.calls = []

    def _urlopen(self, req, timeout=30):
        self.calls.append((req.get_method(), req.full_url))
        if req.get_method() == "POST":
            return _resp(201, [{"id": 42, "updated_at": "2026-11-26T10:00:00Z"}])
        if req.get_method() == "DELETE":
            return _resp(204, None)
        raise AssertionError(f"unexpected request in this test: {req.get_method()} {req.full_url}")

    def test_insert_then_trims_to_just_the_new_row(self):
        with patch("push_to_supabase.urllib.request.urlopen", side_effect=self._urlopen):
            with patch("builtins.open", _fake_open({"ok": True})):
                row = pts.push("fake_path.json")
        self.assertEqual(row["id"], 42)
        methods = [m for m, _ in self.calls]
        self.assertEqual(methods, ["POST", "DELETE"])
        # The trim must keep the just-inserted row and exclude it from deletion.
        self.assertIn("id=neq.42", self.calls[1][1])

    def test_trim_failure_after_a_successful_push_does_not_raise(self):
        def urlopen_trim_fails(req, timeout=30):
            self.calls.append((req.get_method(), req.full_url))
            if req.get_method() == "POST":
                return _resp(201, [{"id": 7, "updated_at": "2026-11-26T10:00:00Z"}])
            raise _http_error(500, {"code": "XXXXX", "message": "some other failure"})

        with patch("push_to_supabase.urllib.request.urlopen", side_effect=urlopen_trim_fails):
            with patch("builtins.open", _fake_open({"ok": True})):
                row = pts.push("fake_path.json")  # must not raise
        self.assertEqual(row["id"], 7)


class TestPushNonTimeoutFailure(_PushTestBase):
    """A plain HTTP failure that is NOT the 57014 statement-timeout code
    must raise immediately, exactly as before this fix -- no wipe, no
    retry, since that failure mode isn't the one this self-heal targets."""

    def test_raises_without_any_delete_or_retry(self):
        calls = []

        def urlopen(req, timeout=30):
            calls.append((req.get_method(), req.full_url))
            raise _http_error(500, {"code": "42501", "message": "permission denied"})

        with patch("push_to_supabase.urllib.request.urlopen", side_effect=urlopen):
            with patch("builtins.open", _fake_open({"ok": True})):
                with self.assertRaises(RuntimeError):
                    pts.push("fake_path.json")
        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0][0], "POST")


class TestPushTimeoutSelfHeal(_PushTestBase):
    """The actual bug this round fixes: INSERT hits 57014. Must look up the
    current newest row, trim everything else (never wiping that one
    fallback row), and retry the insert exactly once."""

    def test_timeout_then_successful_retry_trims_to_new_row(self):
        calls = []

        def urlopen(req, timeout=30):
            calls.append((req.get_method(), req.full_url))
            if req.get_method() == "POST" and len(
                [c for c in calls if c[0] == "POST"]
            ) == 1:
                raise _http_error(500, TIMEOUT_BODY)
            if req.get_method() == "GET":
                return _resp(200, [{"id": 99}])
            if req.get_method() == "DELETE":
                return _resp(204, None)
            if req.get_method() == "POST":
                return _resp(201, [{"id": 123, "updated_at": "2026-11-26T10:05:00Z"}])
            raise AssertionError("unexpected request")

        with patch("push_to_supabase.urllib.request.urlopen", side_effect=urlopen):
            with patch("builtins.open", _fake_open({"ok": True})):
                row = pts.push("fake_path.json")

        self.assertEqual(row["id"], 123)
        methods = [m for m, _ in calls]
        # insert (fails) -> lookup newest -> trim-keep-old -> insert (succeeds) -> trim-keep-new
        self.assertEqual(methods, ["POST", "GET", "DELETE", "POST", "DELETE"])
        # First DELETE (the self-heal) must keep the row that existed BEFORE
        # this run, not the new one -- so a second failure wouldn't wipe the
        # site's last-good data.
        self.assertIn("id=neq.99", calls[2][1])
        # Second DELETE (the normal post-success trim) keeps the new row.
        self.assertIn("id=neq.123", calls[4][1])

    def test_timeout_then_retry_also_fails_leaves_fallback_row_in_place(self):
        calls = []

        def urlopen(req, timeout=30):
            calls.append((req.get_method(), req.full_url))
            if req.get_method() == "GET":
                return _resp(200, [{"id": 55}])
            if req.get_method() == "DELETE":
                return _resp(204, None)
            if req.get_method() == "POST":
                raise _http_error(500, TIMEOUT_BODY)
            raise AssertionError("unexpected request")

        with patch("push_to_supabase.urllib.request.urlopen", side_effect=urlopen):
            with patch("builtins.open", _fake_open({"ok": True})):
                with self.assertRaises(RuntimeError):
                    pts.push("fake_path.json")

        methods = [m for m, _ in calls]
        # insert (fails) -> lookup newest -> trim-keep-old -> insert (fails again) -> stop
        self.assertEqual(methods, ["POST", "GET", "DELETE", "POST"])
        # The one surviving row is the pre-existing one (55), never wiped
        # outright -- the live site still has something to show.
        self.assertIn("id=neq.55", calls[2][1])

    def test_cleanup_lookup_failure_raises_without_a_second_insert_attempt(self):
        calls = []

        def urlopen(req, timeout=30):
            calls.append((req.get_method(), req.full_url))
            if req.get_method() == "POST":
                raise _http_error(500, TIMEOUT_BODY)
            if req.get_method() == "GET":
                raise _http_error(500, {"code": "XXXXX", "message": "down"})
            raise AssertionError("unexpected request")

        with patch("push_to_supabase.urllib.request.urlopen", side_effect=urlopen):
            with patch("builtins.open", _fake_open({"ok": True})):
                with self.assertRaises(RuntimeError):
                    pts.push("fake_path.json")

        methods = [m for m, _ in calls]
        self.assertEqual(methods, ["POST", "GET"])  # never reaches a second insert


class TestStatementTimeoutDetection(unittest.TestCase):
    def test_recognizes_57014_on_http_500(self):
        self.assertTrue(pts._is_statement_timeout(500, json.dumps(TIMEOUT_BODY)))

    def test_other_code_on_http_500_is_not_a_timeout(self):
        self.assertFalse(pts._is_statement_timeout(500, json.dumps({"code": "42501"})))

    def test_57014_text_on_a_different_http_status_is_not_matched(self):
        self.assertFalse(pts._is_statement_timeout(503, json.dumps(TIMEOUT_BODY)))

    def test_unparseable_body_is_not_a_timeout(self):
        self.assertFalse(pts._is_statement_timeout(500, "not json"))


def _fake_open(payload_obj):
    """A drop-in for builtins.open(payload_path) that hands back a context
    manager whose .read()-equivalent (via json.load) yields payload_obj,
    without touching the real filesystem."""
    text = json.dumps(payload_obj)

    def opener(*args, **kwargs):
        return io.StringIO(text)

    return opener


if __name__ == "__main__":
    unittest.main()
