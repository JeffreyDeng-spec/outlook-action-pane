"""Validate account isolation and browser-origin checks on task mutations."""
import http.client
from contextlib import closing
from http.server import ThreadingHTTPServer
import json
import secrets
from threading import Thread
import unittest
from unittest.mock import Mock

from addin_server import make_handler
from windows_worker import FAKE_ACCOUNTS, handler


class ActionRouteTests(unittest.TestCase):
    def setUp(self):
        self.server = ThreadingHTTPServer(("127.0.0.1", 0), None)
        self.token = secrets.token_urlsafe(32)
        self.task = "1" * 32
        self.accounts = {account: account for account in FAKE_ACCOUNTS}
        self.bridge = Mock()
        self.bridge.request.return_value = {"account": FAKE_ACCOUNTS[1], "command": "done", "id": self.task}
        self.headers = {"Origin": f"https://localhost:{self.server.server_port}", "X-Outlook-Triage": "1"}
        self.server.RequestHandlerClass = make_handler(FAKE_ACCOUNTS[0], self.token, self.server.server_port,
                                                      self.accounts, bridge=self.bridge)
        thread = Thread(target=self.server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(self.server.server_close)
        self.addCleanup(thread.join)
        self.addCleanup(self.server.shutdown)

    def request(self, path, headers=None, method="POST", data=b""):
        with closing(http.client.HTTPConnection("localhost", self.server.server_port, timeout=3)) as client:
            client.request(method, path, body=data, headers=headers or {})
            response = client.getresponse()
            return response.status, json.load(response)

    def test_pane_requires_authenticated_path_explicit_account_and_same_origin(self):
        route = f"/p/{self.token}/actions/{self.task}/done"
        selected = "?account=" + FAKE_ACCOUNTS[1]
        self.assertEqual(self.request(route, self.headers)[0], 404)
        self.assertEqual(self.request(route + selected + "&account=" + FAKE_ACCOUNTS[0], self.headers)[0], 404)
        self.assertEqual(self.request(route + selected, {**self.headers, "Origin": "https://example.test"})[0], 403)
        self.assertEqual(self.request(route + selected, self.headers, data=b"x")[0], 403)
        self.assertEqual(self.request(route.replace(self.token, "invalid") + selected, self.headers)[0], 404)
        self.bridge.request.assert_not_called()
        self.assertEqual(self.request(route + selected, self.headers)[0], 200)
        self.bridge.request.assert_called_once_with(f"actions/{self.task}/done", FAKE_ACCOUNTS[1])

    def test_worker_rejects_missing_bearer_unknown_account_and_unsupported_command(self):
        worker = Mock(adapters=self.accounts)
        worker.update_action.return_value = self.bridge.request.return_value
        self.server.RequestHandlerClass = handler(worker, self.token)
        headers = {"Authorization": "Bearer " + self.token}
        route = f"/actions/{self.task}/done"
        self.assertEqual(self.request(route + "?account=" + FAKE_ACCOUNTS[1])[0], 401)
        self.assertEqual(self.request(route + "?account=unbound", headers)[0], 404)
        self.assertEqual(self.request(route.replace("done", "send") + "?account=" + FAKE_ACCOUNTS[1], headers)[0], 404)
        worker.update_action.assert_not_called()
        self.assertEqual(self.request(route + "?account=" + FAKE_ACCOUNTS[1], headers)[0], 200)
        worker.update_action.assert_called_once_with(FAKE_ACCOUNTS[1], self.task, "done")
