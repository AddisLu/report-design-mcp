#!/usr/bin/env python
"""ReportDesign MCP server — stdio, newline-delimited JSON-RPC 2.0, standard library only.

The official MCP Python SDK needs Python >= 3.10; this runs on 3.8.
"""
from __future__ import annotations

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from report_engine import __version__  # noqa: E402
from report_engine import tools  # noqa: E402

PROTOCOL_VERSIONS = ["2025-06-18", "2025-03-26", "2024-11-05"]


class RpcError(Exception):
    def __init__(self, code, message):
        super(RpcError, self).__init__(message)
        self.code = code
        self.message = message


def handle(method, params):
    if method == "initialize":
        asked = params.get("protocolVersion")
        return {
            "protocolVersion": asked if asked in PROTOCOL_VERSIONS else PROTOCOL_VERSIONS[0],
            "capabilities": {"tools": {"listChanged": False}},
            "serverInfo": {"name": "report-design", "version": __version__},
            "instructions": tools.SERVER_INSTRUCTIONS,
        }
    if method == "ping":
        return {}
    if method == "tools/list":
        return {"tools": tools.TOOL_DEFS}
    if method == "tools/call":
        return tools.call_tool(params.get("name"), params.get("arguments") or {})
    raise RpcError(-32601, "Method not found: %s" % method)


def main():
    out = sys.stdout.buffer
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")  # Windows pipes default to cp950
    sys.stdout = sys.stderr  # anything a library prints must not corrupt the protocol stream

    def send(msg):
        out.write(json.dumps(msg, ensure_ascii=False).encode("utf-8") + b"\n")
        out.flush()

    def process(req):
        if not isinstance(req, dict):
            return {"jsonrpc": "2.0", "id": None, "error": {"code": -32600, "message": "Invalid Request"}}
        rid = req.get("id")
        is_notification = "id" not in req
        try:
            result = handle(req.get("method"), req.get("params") or {})
            return None if is_notification else {"jsonrpc": "2.0", "id": rid, "result": result}
        except RpcError as e:
            return None if is_notification else {"jsonrpc": "2.0", "id": rid, "error": {"code": e.code, "message": e.message}}
        except Exception as e:  # never let one request kill the server
            sys.stderr.write("internal error: %r\n" % (e,))
            return None if is_notification else {"jsonrpc": "2.0", "id": rid, "error": {"code": -32603, "message": str(e)}}

    stdin = sys.stdin.buffer
    while True:
        raw = stdin.readline()
        if not raw:
            break
        line = raw.decode("utf-8-sig", errors="replace").strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except ValueError:
            send({"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": "Parse error"}})
            continue
        if isinstance(req, list):
            replies = [r for r in (process(x) for x in req) if r is not None]
            if replies:
                send(replies)
            continue
        if isinstance(req, dict) and req.get("method", "").startswith("notifications/"):
            continue
        reply = process(req)
        if reply is not None:
            send(reply)


if __name__ == "__main__":
    main()
