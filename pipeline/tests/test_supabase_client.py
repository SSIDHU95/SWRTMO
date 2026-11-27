"""
Tests for supabase_client.py's gzip handling, added 27-Nov-2026 while
diagnosing why Supabase egress (outgoing data) had gone well over the
free plan's 5 GB/month limit.

Root cause confirmed directly against the live project: none of this
module's three fetch functions ever asked for a compressed response
(`Accept-Encoding: gzip`), and Python's urllib -- unlike a browser or
`curl --compressed` -- never does this on its own. The same rdps_points
page came back as 172,511 bytes plain vs. 10,189 bytes gzip-compressed
(about 17x), and that table alone was being re-fetched in full up to
twice a day.

These tests mock urllib.request.urlopen entirely (no real network calls)
and check that (a) every outgoing request now asks for gzip, and (b) a
gzip-compressed response is transparently decompressed back into the
same rows a plain response would have produced -- the two code paths
must be indistinguishable to every calling source adapter.
"""
import gzip
import json
import os
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import supabase_client as sc


def _resp(body_bytes, gzipped=False):
    """A fake context-manager response, like what urlopen(...) as resp
    gives. `headers.get(...)` mimics http.client.HTTPMessage's API."""
    m = MagicMock()
    m.__enter__.return_value = m
    m.__exit__.return_value = False
    sent = gzip.compress(body_bytes) if gzipped else body_bytes
    m.read.return_value = sent
    m.headers = MagicMock()
    m.headers.get.side_effect = lambda k, default="": (
        "gzip" if gzipped and k == "Content-Encoding" else default
    )
    return m


class _EnvBase(unittest.TestCase):
    def setUp(self):
        patcher = patch.dict(os.environ, {
            "SUPABASE_URL": "https://x.supabase.co",
            "SUPABASE_ANON_KEY": "test-anon-key",
        })
        patcher.start()
        self.addCleanup(patcher.stop)


class TestFetchLatestPayloadGzip(_EnvBase):
    def test_request_asks_for_gzip(self):
        captured = {}

        def fake_urlopen(req, timeout=30):
            captured["headers"] = dict(req.headers)
            body = json.dumps([{"payload": {"a": 1}, "updated_at": "t"}]).encode()
            return _resp(body, gzipped=False)

        with patch("urllib.request.urlopen", side_effect=fake_urlopen):
            payload, updated_at = sc.fetch_latest_payload("some_table")
        # urllib title-cases header keys it's given (Accept-encoding), so
        # compare case-insensitively rather than assuming exact casing.
        headers_lower = {k.lower(): v for k, v in captured["headers"].items()}
        self.assertEqual(headers_lower.get("accept-encoding"), "gzip")
        self.assertEqual(payload, {"a": 1})

    def test_gzip_compressed_response_decodes_to_same_result_as_plain(self):
        rows = [{"payload": {"x": [1, 2, 3]}, "updated_at": "2026-11-27T00:00:00Z"}]
        body = json.dumps(rows).encode()

        with patch("urllib.request.urlopen", return_value=_resp(body, gzipped=False)):
            plain_result = sc.fetch_latest_payload("some_table")
        with patch("urllib.request.urlopen", return_value=_resp(body, gzipped=True)):
            gzip_result = sc.fetch_latest_payload("some_table")

        self.assertEqual(plain_result, gzip_result)
        self.assertEqual(gzip_result, ({"x": [1, 2, 3]}, "2026-11-27T00:00:00Z"))


class TestFetchAllPayloadsGzip(_EnvBase):
    def test_gzip_compressed_response_decodes_correctly(self):
        rows = [
            {"payload": {"n": 1}, "updated_at": "2026-11-25T00:00:00Z"},
            {"payload": {"n": 2}, "updated_at": "2026-11-26T00:00:00Z"},
        ]
        body = json.dumps(rows).encode()
        with patch("urllib.request.urlopen", return_value=_resp(body, gzipped=True)):
            result = sc.fetch_all_payloads("some_table")
        self.assertEqual(result, [
            ({"n": 1}, "2026-11-25T00:00:00Z"),
            ({"n": 2}, "2026-11-26T00:00:00Z"),
        ])


class TestFetchRowsGzip(_EnvBase):
    def test_paginated_gzip_responses_concatenate_correctly(self):
        """fetch_rows() pages through a table; this confirms gzip
        decoding works correctly across multiple pages, not just a
        single-page response, and that pagination still stops at the
        real end of the table (a short final page) exactly as before."""
        page1 = [{"id": i} for i in range(1000)]
        page2 = [{"id": i} for i in range(1000, 1500)]

        responses = [
            _resp(json.dumps(page1).encode(), gzipped=True),
            _resp(json.dumps(page2).encode(), gzipped=True),
        ]

        def fake_urlopen(req, timeout=30):
            return responses.pop(0)

        with patch("urllib.request.urlopen", side_effect=fake_urlopen):
            rows = sc.fetch_rows("rdps_points", select="id", page_size=1000)

        self.assertEqual(len(rows), 1500)
        self.assertEqual(rows[0]["id"], 0)
        self.assertEqual(rows[-1]["id"], 1499)

    def test_every_page_request_asks_for_gzip(self):
        captured_headers = []

        def fake_urlopen(req, timeout=30):
            captured_headers.append({k.lower(): v for k, v in dict(req.headers).items()})
            return _resp(json.dumps([]).encode(), gzipped=False)

        with patch("urllib.request.urlopen", side_effect=fake_urlopen):
            sc.fetch_rows("rdps_points", select="id")

        self.assertTrue(captured_headers)
        for h in captured_headers:
            self.assertEqual(h.get("accept-encoding"), "gzip")


class TestMaybeGunzipHelper(_EnvBase):
    def test_plain_body_passed_through_unchanged(self):
        resp = _resp(b'{"a":1}', gzipped=False)
        self.assertEqual(sc._maybe_gunzip(resp, b'{"a":1}'), b'{"a":1}')

    def test_gzip_body_is_decompressed(self):
        raw = b'{"a":1}'
        resp = _resp(raw, gzipped=True)
        compressed = gzip.compress(raw)
        self.assertEqual(sc._maybe_gunzip(resp, compressed), raw)


if __name__ == "__main__":
    unittest.main()
