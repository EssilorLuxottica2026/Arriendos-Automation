import unittest

from lease_accounting.web import (
    _bot_lucy_finish_event,
    _bot_lucy_lock,
    _bot_lucy_state,
    app,
)


class BotLucyWebTests(unittest.TestCase):
    def setUp(self):
        with _bot_lucy_lock:
            self.previous_state = {
                "status": _bot_lucy_state["status"],
                "error": _bot_lucy_state["error"],
                "logs": list(_bot_lucy_state["logs"]),
            }
            _bot_lucy_state.update(status="idle", error=None, logs=[])
        _bot_lucy_finish_event.clear()
        self.client = app.test_client()

    def tearDown(self):
        with _bot_lucy_lock:
            _bot_lucy_state.update(self.previous_state)
        _bot_lucy_finish_event.clear()

    def test_status_returns_only_logs_after_requested_cursor(self):
        with _bot_lucy_lock:
            _bot_lucy_state.update(
                status="running",
                logs=["first", "second", "third"],
            )

        response = self.client.get("/bot-lucy-status?after=1")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.get_json(),
            {
                "status": "running",
                "message": None,
                "logs": ["second", "third"],
                "total": 3,
            },
        )

    def test_finish_request_releases_only_a_held_bot(self):
        with _bot_lucy_lock:
            _bot_lucy_state["status"] = "holding"

        response = self.client.post("/finish-bot-lucy")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["status"], "finishing")
        self.assertTrue(_bot_lucy_finish_event.is_set())

    def test_finish_request_does_not_stop_a_running_bot(self):
        with _bot_lucy_lock:
            _bot_lucy_state["status"] = "running"

        response = self.client.post("/finish-bot-lucy")

        self.assertEqual(response.status_code, 409)
        self.assertEqual(response.get_json()["status"], "running")
        self.assertFalse(_bot_lucy_finish_event.is_set())


if __name__ == "__main__":
    unittest.main()
