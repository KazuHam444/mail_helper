import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch, Mock
from datetime import datetime, timezone

import hr_digest as app


def mail(subject="面试邀请", body="感谢应聘，请确认明天面试时间。", mid="m1"):
    return dict(id=mid, subject=subject, body=body, sender="hr@company.example",
                sender_name="HR", received="2026-11-03T08:00:00Z",
                url="https://outlook.live.com/mail/0/inbox/id/test")


class RulesTests(unittest.TestCase):
    def setUp(self):
        self.env = patch.dict(os.environ, {"ALLOW_SENDERS": "", "BLOCK_SENDERS": "", "AI_API_KEY": "", "AI_MODEL": ""})
        self.env.start()
        self.addCleanup(self.env.stop)

    def test_interview_and_form_and_reschedule(self):
        for subject, body in [("面试邀请", "感谢应聘，请确认面试时间。"),
                              ("请完成候选人信息表填写", "感谢投递，请填写候选人登记表。"),
                              ("请完成候选人信息表", "请按要求提交信息。"),
                              ("Interview rescheduled", "Please confirm your interview."),
                              ("请补充应聘材料", "请提交您的作品和简历。")]:
            self.assertTrue(app.rules(mail(subject, body))[0]["notify"], subject)

    def test_ads_and_receipts_and_rejection(self):
        for subject, body in [("为你推荐高薪职位", "请参加面试课程"),
                              ("投递成功", "已收到您的简历。"),
                              ("Interview tips newsletter", "Interview invitation examples"),
                              ("申请进展", "很遗憾我们无法继续此次申请。"),
                              ("订单发货通知", "请填写收货信息")]:
            self.assertFalse(app.rules(mail(subject, body))[0]["notify"], subject)

    def test_obvious_fee_scam(self):
        self.assertFalse(app.rules(mail("面试邀请", "面试前请支付保证金。"))[0]["notify"])

    def test_domain_matching_not_substring(self):
        self.assertTrue(app.sender_matches("hr@company.com", ["company.com"]))
        self.assertFalse(app.sender_matches("hr@evilcompany.com", ["company.com"]))

    def test_block_overrides_allow(self):
        with patch.dict(os.environ, {"ALLOW_SENDERS": "company.example", "BLOCK_SENDERS": "hr@company.example"}):
            self.assertFalse(app.rules(mail())[0]["notify"])

    def test_ai_outage_fallback_remains_retryable(self):
        with patch.dict(os.environ, {"AI_API_KEY": "fake", "AI_MODEL": "fake"}), patch.object(app, "ai_classify", side_effect=TimeoutError):
            decision, final = app.classify(mail())
            self.assertTrue(decision["notify"])
            self.assertFalse(final)

    def test_non_outlook_url_not_forwarded(self):
        m = mail()
        m["url"] = "https://outlook.live.com.attacker.example/"
        self.assertNotIn("attacker", app.item_text(m, app.result(True)))


class StateTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.state = app.State(Path(self.temp.name) / "state.db")
        self.addCleanup(self.state.db.close)

    def test_fallback_can_be_reclassified(self):
        self.state.save(mail(), app.result(True), False)
        self.assertFalse(self.state.seen("m1"))
        self.state.save(mail(), app.result(False), True)
        self.assertTrue(self.state.seen("m1"))
        self.assertEqual(self.state.db.execute("SELECT count(*) FROM messages").fetchone()[0], 1)

    def test_outbox_retry_per_channel_and_deduplication(self):
        self.state.enqueue("m1", "test", "body", ["desktop", "telegram"])
        self.state.enqueue("m1", "test", "body", ["desktop", "telegram"])
        with patch.object(app, "channels", return_value=["desktop", "telegram"]), patch.object(app, "desktop"), patch.object(app, "telegram", side_effect=TimeoutError):
            self.assertFalse(app.deliver(self.state))
        with patch.object(app, "channels", return_value=["desktop", "telegram"]), patch.object(app, "desktop") as desk, patch.object(app, "telegram") as tg:
            self.assertTrue(app.deliver(self.state))
            desk.assert_not_called()
            tg.assert_called_once()

    def test_disabled_channel_does_not_send_pending(self):
        self.state.enqueue("m1", "test", "body", ["telegram"])
        with patch.object(app, "channels", return_value=[]), patch.object(app, "telegram") as tg:
            self.assertTrue(app.deliver(self.state))
            tg.assert_not_called()
        self.assertEqual(self.state.db.execute("SELECT sent FROM outbox").fetchone()[0], 0)

    def test_local_report_available_before_digest_time(self):
        with patch.object(app, "DATA", Path(self.temp.name)), patch.object(app, "channels", return_value=["telegram"]):
            self.state.save(mail(), app.result(True))
            app.make_digest(self.state, datetime(2026, 11, 3, 9, 0, tzinfo=timezone.utc))
            self.assertTrue((Path(self.temp.name) / "reports/2026-11-03.txt").exists())
            self.assertEqual(self.state.db.execute("SELECT count(*) FROM outbox").fetchone()[0], 0)

    def test_daily_digest_filters_day_and_only_once(self):
        with patch.object(app, "DATA", Path(self.temp.name)), patch.object(app, "channels", return_value=["telegram"]):
            self.state.save(mail(), app.result(True, "面试", "确认时间"))
            old = mail(mid="old")
            old["received"] = "2026-11-02T08:00:00Z"
            self.state.save(old, app.result(True))
            stamp = datetime(2026, 11, 3, 12, 0, tzinfo=timezone.utc)
            app.make_digest(self.state, stamp)
            app.make_digest(self.state, stamp)
            report = (Path(self.temp.name) / "reports/2026-11-03.txt").read_text(encoding="utf-8")
            self.assertIn("1 封", report)
            self.assertEqual(self.state.db.execute("SELECT count(*) FROM outbox").fetchone()[0], 1)

    def test_mid_scan_failure_does_not_advance_watermark(self):
        def broken(*args):
            yield mail()
            raise RuntimeError("page 2 failed")
        with patch.object(app, "token", return_value="fake"), patch.object(app, "fetch_messages", broken), patch.object(app, "channels", return_value=[]):
            with self.assertRaises(RuntimeError):
                app.run(self.state)
        self.assertIsNone(self.state.get("watermark"))
        self.assertTrue(self.state.seen("m1"))


class GraphTests(unittest.TestCase):
    def test_pagination_without_reapplying_params(self):
        raw = {"id": "id1", "receivedDateTime": "2026-11-03T00:00:00Z", "body": {"content": "text"}}
        pages = [{"value": [raw], "@odata.nextLink": app.GRAPH + "/next"}, {"value": []}]
        with patch.dict(os.environ, {"MAIL_FOLDERS": "inbox"}), patch.object(app, "graph_get", side_effect=pages) as get:
            self.assertEqual(len(list(app.fetch_messages("fake", "start", "end"))), 1)
            self.assertIsNone(get.call_args_list[1].args[2])

    def test_no_token_to_other_hosts(self):
        with self.assertRaises(RuntimeError):
            app.graph_get("https://graph.microsoft.com.evil.example/messages", "secret")


class AITests(unittest.TestCase):
    def setUp(self):
        self.env = patch.dict(os.environ, {"AI_BASE_URL": "https://ai.example/v1", "AI_API_KEY": "test-key",
                                          "AI_MODEL": "test-model", "AI_MIN_CONFIDENCE": "0.8"})
        self.env.start()
        self.addCleanup(self.env.stop)

    def call(self, data):
        response = Mock(status_code=200)
        response.json.return_value = {"choices": [{"message": {"content": json.dumps(data)}}]}
        with patch("requests.post", return_value=response) as post:
            decision = app.ai_classify(mail())
            self.assertEqual(post.call_args.args[0], "https://ai.example/v1/chat/completions")
            self.assertFalse(post.call_args.kwargs["allow_redirects"])
            return decision

    def test_valid_action_and_low_confidence(self):
        data = dict(notify=True, confidence=0.95, category="面试", summary="请确认面试时间", deadline="")
        self.assertTrue(self.call(data)["notify"])
        data["confidence"] = 0.5
        self.assertFalse(self.call(data)["notify"])

    def test_invalid_boolean_rejected(self):
        with self.assertRaises(ValueError):
            self.call(dict(notify="true", confidence=0.95, category="面试", summary="确认", deadline=""))

    def test_nonfinite_confidence_rejected(self):
        with self.assertRaises(ValueError):
            self.call(dict(notify=True, confidence=float("nan"), category="面试", summary="确认", deadline=""))

    def test_insecure_service_rejected(self):
        with patch.dict(os.environ, {"AI_BASE_URL": "http://ai.example/v1"}):
            with self.assertRaises(RuntimeError):
                app.ai_classify(mail())

    def test_dashscope_key_defaults_and_non_thinking_json(self):
        with patch.dict(os.environ, {"AI_BASE_URL": "", "AI_API_KEY": "", "AI_MODEL": "", "DASHSCOPE_API_KEY": "fake-dash-key"}):
            self.assertTrue(app.ai_enabled())
            self.assertEqual(app.ai_settings(), ("https://dashscope.aliyuncs.com/compatible-mode/v1", "fake-dash-key", "qwen-plus"))
            response = Mock(status_code=200)
            response.json.return_value = {"choices": [{"message": {"content": json.dumps(dict(
                notify=True, confidence=0.95, category="面试", summary="请确认面试", deadline=""))}}]}
            with patch("requests.post", return_value=response) as post:
                self.assertTrue(app.ai_classify(mail())["notify"])
                self.assertEqual(post.call_args.kwargs["headers"]["Authorization"], "Bearer fake-dash-key")
                self.assertFalse(post.call_args.kwargs["json"]["enable_thinking"])

    def test_dashscope_key_never_sent_to_old_openai_default(self):
        with patch.dict(os.environ, {"AI_API_KEY": "", "DASHSCOPE_API_KEY": "fake-dash-key", "AI_BASE_URL": "https://api.openai.com/v1"}):
            with patch("requests.post") as post, self.assertRaises(RuntimeError):
                app.ai_classify(mail())
            post.assert_not_called()


if __name__ == "__main__":
    unittest.main()
