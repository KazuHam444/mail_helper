import email
import contextlib
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
                url="https://mail.qq.com/")


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
                              ("请补充应聘材料", "请提交您的作品和简历。"),
                              ("Next stage - video submission",
                               "We are pleased to proceed with your application to the next stage. "
                               "Please record a 5-minute video and submit your video within 48 hours. "
                               "The link will be closed by Wednesday October 28, 11:59pm HKT. Internship Program.")]:
            decision, candidate = app.rules(mail(subject, body))
            self.assertTrue(candidate, subject)
            self.assertTrue(decision["notify"], subject)

    def test_ads_and_receipts_and_rejection(self):
        for subject, body in [("为你推荐高薪职位", "请参加面试课程"),
                              ("投递成功", "已收到您的简历。"),
                              ("Interview tips newsletter", "Interview invitation examples"),
                              ("申请进展", "很遗憾我们无法继续此次申请。"),
                              ("订单发货通知", "请填写收货信息")]:
            m = mail(subject, body)
            m["sender"] = "noreply@jobboard.example"
            m["sender_name"] = "Job Board"
            self.assertFalse(app.rules(m)[0]["notify"], subject)

    def test_any_hr_sender_notifies_including_receipt_and_reject(self):
        for subject, body in [
            ("Application received", "Thank you for applying. We have received your application."),
            ("Update", "We regret to inform you that your application was unsuccessful."),
            ("Quick note", "Just checking in regarding your application status."),
        ]:
            m = mail(subject, body)
            m["sender"] = "HKCareers@polymercapital.com"
            m["sender_name"] = "HK Careers"
            decision, candidate = app.rules(m)
            self.assertTrue(candidate, subject)
            self.assertTrue(decision["notify"], subject)

    def test_obvious_fee_scam(self):
        self.assertFalse(app.rules(mail("面试邀请", "面试前请支付保证金。"))[0]["notify"])

    def test_student_card_nol_request_from_named_person(self):
        m = mail("Re: Internship", "Dear Zijun, thanks and may you please send me your student card, NOL? "
                 "Also, if there is an opportunity to take up the internship, do you need to defer your graduation?")
        m.update(sender="alice@company.example", sender_name="Alice")
        decision, candidate = app.rules(m)
        self.assertTrue(candidate)
        self.assertTrue(decision["notify"])

    def test_linkedin_recommendations_and_university_digest_never_call_ai(self):
        samples = [
            ("金融财务职位", "职位订阅: 金融财务\nFinance Analyst (12 months contract)\n快速申请", "jobalerts-noreply@linkedin.com", "领英职位订阅"),
            ("[Ext] 您可能适合BinanceFinance Analyst (12 months contract)职位 - 积极招聘", "职位订阅: Accountant\n正在热招\n快速申请", "jobalerts-noreply@linkedin.com", "领英职位订阅"),
            ("CityUHK Announcement Portal (CAP)", "Dear SHEN Zijun,\nNew posting(s) in CityUHK Announcement Portal (CAP) summarized below:\nCareer Preparation Workshop | Zoom\nRolling Application for Hall. You are invited to join.\nEmail me a copy", "cap@example.edu", "University Portal"),
        ]
        for subject, body, sender, name in samples:
            with self.subTest(subject=subject), patch.dict(os.environ, {"AI_API_KEY": "fake", "AI_MODEL": "fake"}), patch.object(app, "ai_classify") as ai:
                m = mail(subject, body)
                m.update(sender=sender, sender_name=name)
                decision, final = app.classify(m)
                self.assertFalse(decision["notify"])
                self.assertTrue(final)
                ai.assert_not_called()

    def test_broadcast_exclusion_precedes_hr_and_allowlist(self):
        with patch.dict(os.environ, {"ALLOW_SENDERS": "linkedin.com,company.example"}):
            for sender in ("hr@company.example", "jobalerts-noreply@linkedin.com"):
                m = mail("职位订阅: 面试岗位推荐", "请申请这些职位。")
                m["sender"] = sender
                self.assertFalse(app.rules(m)[0]["notify"])

    def test_forwarded_broadcast_still_excluded(self):
        m = mail("FW: Updates", "From: 领英职位订阅 <jobalerts-noreply@linkedin.com>\n职位订阅: Accountant\n快速申请")
        self.assertFalse(app.rules(m)[0]["notify"])

    def test_linkedin_and_university_personal_interviews_not_domain_blocked(self):
        for sender in ("messages-noreply@linkedin.com", "alice@university.example", "hkcareers@company.example"):
            m = mail("Interview confirmation", "Your application was shortlisted. Please confirm your interview time.")
            m.update(sender=sender, sender_name="Alice")
            self.assertTrue(app.rules(m)[0]["notify"])

    def test_jobalerts_or_hr_in_domain_does_not_make_hr_sender(self):
        for sender in ("jobalerts-noreply@linkedin.com", "noreply@careers.example", "myjob@company.example"):
            self.assertFalse(app.looks_like_hr(dict(sender=sender, sender_name="Updates")))

    def test_synthetic_dataset(self):
        cases = json.loads((app.ROOT / 'tests/hr_eval_cases.json').read_text(encoding='utf-8'))
        for case in cases:
            with self.subTest(case=case['id']):
                self.assertEqual(app.rules(case)[0]['notify'], case['expected'])

    def test_personal_assessment_workshop_is_not_public_training(self):
        m = mail('Assessment workshop invitation', 'Your application has progressed. You are invited to an assessment workshop. Please confirm your attendance.')
        m.update(sender='alice@company.example', sender_name='Alice')
        self.assertTrue(app.rules(m)[0]['notify'])

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

    def test_ai_cannot_drop_explicit_personal_application_receipt(self):
        m = mail('Application received', 'We have received your application.')
        with patch.dict(os.environ, {"AI_API_KEY": "fake", "AI_MODEL": "fake"}), patch.object(app, 'ai_classify', return_value=app.result(False, source='ai')):
            decision, final = app.classify(m)
            self.assertTrue(decision['notify'])
            self.assertTrue(final)

    def test_ai_cannot_drop_direct_cv_and_start_date_request(self):
        m = mail('Re: Analyst role', 'Could you send us your CV and earliest start date for the analyst position?')
        m.update(sender='alice@company.example', sender_name='Alice')
        with patch.dict(os.environ, {"AI_API_KEY": "fake", "AI_MODEL": "fake"}), patch.object(app, 'ai_classify', return_value=app.result(False, source='ai')):
            self.assertTrue(app.classify(m)[0]['notify'])

    def test_non_qq_url_not_forwarded(self):
        m = mail()
        m["url"] = "https://mail.qq.com.attacker.example/"
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

    def test_upgrade_cancels_pending_broadcast_but_keeps_real_hr(self):
        advert = mail("职位订阅", "职位订阅: Accountant", "advert")
        self.state.save(advert, app.result(True))
        self.state.save(mail(mid="real"), app.result(True))
        self.state.enqueue("mail:advert:0", "old", "advert", ["desktop"])
        self.state.enqueue("mail:real:0", "real", "interview", ["desktop"])
        self.state.enqueue("digest:2026-11-03:0", "old digest", "stale", ["desktop"])
        app.migrate_broadcast_filter(self.state)
        app.migrate_broadcast_filter(self.state)
        keys = [row[0] for row in self.state.db.execute("SELECT key FROM outbox")]
        self.assertEqual(keys, ["mail:real:0:desktop"])
        decision = self.state.db.execute("SELECT decision FROM messages WHERE id='advert'").fetchone()[0]
        self.assertFalse(json.loads(decision)["notify"])

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

    def test_catchup_yesterday_after_1230_once(self):
        with patch.dict(os.environ, {"CATCHUP_DIGEST_TIME": "12:30"}), patch.object(app, "channels", return_value=["desktop"]):
            yesterday = mail(mid="y1")
            yesterday["received"] = "2026-11-03T08:00:00Z"
            self.state.save(yesterday, app.result(True, "面试", "确认时间"))
            before = datetime(2026, 11, 4, 4, 0, tzinfo=timezone.utc)  # 北京 12:00
            app.make_catchup(self.state, before)
            self.assertEqual(self.state.db.execute("SELECT count(*) FROM outbox").fetchone()[0], 0)
            after = datetime(2026, 11, 4, 4, 40, tzinfo=timezone.utc)  # 北京 12:40
            app.make_catchup(self.state, after)
            app.make_catchup(self.state, after)
            rows = self.state.db.execute("SELECT key,title,body FROM outbox").fetchall()
            self.assertEqual(len(rows), 1)
            self.assertIn("catchup:2026-11-03", rows[0][0])
            self.assertEqual(rows[0][1], "猪猪猪猪！！！")
            self.assertIn("昨日招聘摘要", rows[0][2])

    def test_digest_email_enqueued_once_without_chunk_split(self):
        with patch.dict(os.environ, {"EMAIL_TO": "a@qq.com,b@example.com", "DESKTOP_NOTIFY": "false",
                                     "TELEGRAM_BOT_TOKEN": "", "TELEGRAM_CHAT_ID": ""}):
            long_body = "摘要内容" * 1000
            app.enqueue_text(self.state, "digest:2026-11-04", "猪猪猪猪！！！", long_body, app.channels("digest"))
            rows = self.state.db.execute("SELECT key,channel,body FROM outbox").fetchall()
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0][1], "email")
            self.assertEqual(len(rows[0][2]), len(long_body))
            self.assertEqual(app.email_recipients(), ["a@qq.com", "b@example.com"])
            self.assertNotIn("email", app.channels("alert"))

    def test_mid_scan_failure_does_not_advance_watermark(self):
        def broken(*args):
            yield mail()
            raise RuntimeError("page 2 failed")
        with patch.object(app, "fetch_messages", broken), patch.object(app, "channels", return_value=[]):
            with self.assertRaises(RuntimeError):
                app.run(self.state)
        self.assertIsNone(self.state.get("watermark"))
        self.assertTrue(self.state.seen("m1"))

    def test_match_delivered_before_later_scan_failure(self):
        def broken(*args):
            yield mail()
            self.assertEqual(self.state.db.execute("SELECT sent FROM outbox").fetchone()[0], 1)
            raise RuntimeError("later fetch failed")
        with patch.object(app, "fetch_messages", broken), patch.object(app, "classify", return_value=(app.result(True), True)), patch.object(app, "channels", return_value=["desktop"]), patch.object(app, "desktop") as desk, patch.object(app, "append_notify_log"):
            with self.assertRaises(RuntimeError):
                app.run(self.state)
            desk.assert_called_once()

    def test_pending_delivered_even_when_mailbox_unavailable(self):
        self.state.enqueue("pending", "test", "body", ["desktop"])
        with patch.object(app, "fetch_messages", side_effect=RuntimeError("offline")), patch.object(app, "channels", return_value=["desktop"]), patch.object(app, "desktop") as desk:
            with self.assertRaises(RuntimeError):
                app.run(self.state)
            desk.assert_called_once()
            self.assertEqual(self.state.db.execute("SELECT sent FROM outbox").fetchone()[0], 1)


class ImapTests(unittest.TestCase):
    def test_delayed_mail_uses_internal_arrival_not_old_date(self):
        client = Mock()
        client.select.return_value = ("OK", [b"1"])
        raw = b"From: hr@example.com\r\nDate: Sun, 1 Nov 2026 10:00:00 +0000\r\nMessage-ID: <delayed>\r\nSubject: Interview\r\n\r\nPlease confirm."
        client.uid.side_effect = [("OK", [b"123"]), ("OK", [(b'1 (UID 123 INTERNALDATE "04-Nov-2026 16:00:00 +0800" BODY[] {100}', raw), b')'])]
        with patch.object(app, "imap_connection", return_value=contextlib.nullcontext(client)), patch.dict(os.environ, {"MAIL_FOLDERS": "INBOX"}):
            messages = list(app.fetch_messages("2026-11-04T07:00:00Z", "2026-11-04T09:00:00Z"))
        self.assertEqual(len(messages), 1)
        self.assertEqual(messages[0]["received"], "2026-11-04T08:00:00Z")
        self.assertEqual(client.uid.call_args_list[0].args[0], "search")
        self.assertEqual(client.uid.call_args_list[1].args[:2], ("fetch", b"123"))


class DesktopTests(unittest.TestCase):
    def test_missing_notification_history_signal_is_failure(self):
        with patch.dict(os.environ, {"DESKTOP_STYLE": "window"}), patch.object(app.subprocess, "run", return_value=Mock(returncode=0, stdout=b"")):
            with self.assertRaises(RuntimeError):
                app.desktop("test", "body")

    def test_notification_history_signal_is_accepted(self):
        with patch.dict(os.environ, {"DESKTOP_STYLE": "window"}), patch.object(app.subprocess, "run", return_value=Mock(returncode=0, stdout=b"HR_TOAST_IN_HISTORY\r\n")):
            app.desktop("test", "body")

    def test_old_window_setting_cannot_create_foreground_window(self):
        with patch.dict(os.environ, {"DESKTOP_STYLE": "window", "DESKTOP_BANNER": "false"}), patch.object(app.subprocess, "run", return_value=Mock(returncode=0, stdout=b"HR_TOAST_IN_HISTORY")) as run:
            app.desktop("test", "body")
            script = run.call_args.args[0][-1]
            self.assertNotIn("ShowDialog", script)
            self.assertNotIn("TopMost", script)
            self.assertIn("SuppressPopup", script)
            self.assertEqual(run.call_args.kwargs["env"]["HR_NOTIFY_BANNER"], "false")


class MoreImapTests(unittest.TestCase):
    def test_stable_id_prefers_message_id(self):
        msg = email.message_from_string("Message-ID: <abc@example.com>\nSubject: hi\n\nbody")
        first = app.stable_mail_id("INBOX", "1", msg, "2026-11-03T00:00:00Z")
        second = app.stable_mail_id("INBOX", "99", msg, "2026-11-04T00:00:00Z")
        self.assertEqual(first, second)
        self.assertTrue(first.startswith("mid:"))

    def test_forwarded_sender_restored_from_body(self):
        body = "---------- Forwarded message ----------\nFrom: HR Team <hr@company.example>\n\n请确认明天面试。"
        msg = email.message_from_string(
            "From: me@hotmail.com\nSubject: FW: 面试邀请\n\n" + body)
        with patch.dict(os.environ, {"QQ_EMAIL": "me@qq.com", "FORWARD_SENDERS": "hotmail.com"}):
            addr, name = app.resolve_sender(msg, body)
        self.assertEqual(addr, "hr@company.example")
        self.assertIn("HR", name)

    def test_html_body_extracted(self):
        raw = (
            "From: hr@x.com\r\nSubject: test\r\nMIME-Version: 1.0\r\n"
            "Content-Type: text/html; charset=utf-8\r\n"
            "Content-Transfer-Encoding: 8bit\r\n\r\n"
            "<p>请确认面试时间</p>"
        ).encode("utf-8")
        msg = email.message_from_bytes(raw)
        self.assertIn("请确认面试时间", app.extract_body(msg))

    def test_imap_settings_require_auth_code(self):
        with patch.dict(os.environ, {"QQ_EMAIL": "a@qq.com", "QQ_AUTH_CODE": ""}, clear=False):
            with self.assertRaises(RuntimeError):
                app.imap_settings()


class AITests(unittest.TestCase):
    def setUp(self):
        self.env = patch.dict(os.environ, {"AI_BASE_URL": "https://ai.example/v1", "AI_API_KEY": "test-key",
                                          "AI_MODEL": "test-model", "AI_REJECT_CONFIDENCE": "0.85"})
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

    def test_valid_action_and_uncertain_reject_still_notifies(self):
        data = dict(notify=True, confidence=0.4, category="面试", summary="请确认面试时间", deadline="")
        self.assertTrue(self.call(data)["notify"])
        # AI 否定但把握不足 -> 仍提醒（宁可错报）
        data = dict(notify=False, confidence=0.5, category="ignore", summary="", deadline="")
        self.assertTrue(self.call(data)["notify"])
        # AI 明确否定 -> 静默
        data = dict(notify=False, confidence=0.95, category="ignore", summary="", deadline="")
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
