"""Gemini request contract tests. These never contact an API or use a real key."""

from types import SimpleNamespace
from unittest.mock import MagicMock, patch
import unittest

from google.genai import errors, types

import ai_service


FAKE_KEY = "test-only-private-key"


def response(text="Generated lesson", *, block_reason=None, finish_reason="STOP"):
    return SimpleNamespace(
        text=text,
        prompt_feedback=SimpleNamespace(block_reason=block_reason),
        candidates=[SimpleNamespace(finish_reason=finish_reason)],
    )


def api_error(code, *, reason=None):
    body = {"error": {"code": code, "message": f"Secret {FAKE_KEY}"}}
    if reason:
        body["error"]["details"] = [{"reason": reason}]
    return errors.APIError(code, body)


class GenerateTextTests(unittest.TestCase):
    def setUp(self):
        self.client = MagicMock()
        self.client.models.generate_content.return_value = response()
        self.client_patch = patch.object(ai_service.genai, "Client", return_value=self.client)
        self.factory = self.client_patch.start()
        self.addCleanup(self.client_patch.stop)
        self.sleep_patch = patch.object(ai_service.time, "sleep")
        self.sleep = self.sleep_patch.start()
        self.addCleanup(self.sleep_patch.stop)

    def assert_safe_error(self, expected, call=None):
        with self.assertRaises(ai_service.AIError) as caught:
            (call or (lambda: ai_service.generate_text("Lesson", FAKE_KEY)))()
        self.assertEqual(caught.exception.code, expected)
        self.assertEqual(str(caught.exception), expected)
        self.assertNotIn(FAKE_KEY, str(caught.exception))
        self.assertIsNone(caught.exception.__cause__)
        return caught.exception

    def test_text_is_forwarded_and_client_closed(self):
        result = ai_service.generate_text("  Сабақ жоспары  ", FAKE_KEY, model="custom-model")
        self.assertEqual(result, "Generated lesson")
        kwargs = self.client.models.generate_content.call_args.kwargs
        self.assertEqual(kwargs["model"], "custom-model")
        self.assertEqual(kwargs["contents"][0].role, "user")
        self.assertEqual(kwargs["contents"][0].parts[0].text, "  Сабақ жоспары  ")
        self.assertEqual(kwargs["config"].response_mime_type, "text/plain")
        self.assertEqual(self.factory.call_args.kwargs["api_key"], FAKE_KEY)
        options = self.factory.call_args.kwargs["http_options"]
        self.assertEqual(options.timeout, 60_000)
        self.assertEqual(options.retry_options.attempts, 1)
        self.client.close.assert_called_once_with()

    def test_json_mode_and_attachments_are_forwarded(self):
        self.client.models.generate_content.return_value = response(' {"topic": "ҚМЖ"} ')
        result = ai_service.generate_text(
            "Read files", FAKE_KEY, json_mode=True,
            attachments=[("application/pdf", b"%PDF-1.4 demo"), ("image/png", b"demo image")],
        )
        self.assertEqual(result, '{"topic": "ҚМЖ"}')
        kwargs = self.client.models.generate_content.call_args.kwargs
        self.assertEqual(kwargs["config"].response_mime_type, "application/json")
        parts = kwargs["contents"][0].parts
        self.assertEqual(len(parts), 3)
        self.assertEqual(parts[1].inline_data.mime_type, "application/pdf")
        self.assertEqual(parts[1].inline_data.data, b"%PDF-1.4 demo")
        self.assertEqual(parts[2].inline_data.mime_type, "image/png")
        self.assertEqual(parts[2].inline_data.data, b"demo image")
        self.client.close.assert_called_once_with()

    def test_missing_key_does_not_construct_client(self):
        self.assert_safe_error("missing_key", lambda: ai_service.generate_text("Lesson", "  "))
        self.factory.assert_not_called()

    def test_invalid_input_does_not_construct_client(self):
        self.assert_safe_error("request", lambda: ai_service.generate_text(" ", FAKE_KEY))
        self.assert_safe_error("request", lambda: ai_service.generate_text("Lesson", FAKE_KEY, attachments=[("application/pdf", b"")]))
        self.factory.assert_not_called()

    def test_auth_and_quota_are_sanitized_without_retries(self):
        for status, expected in [(401, "auth"), (403, "auth"), (429, "quota"), (404, "unavailable")]:
            with self.subTest(status=status):
                self.client.reset_mock()
                self.client.models.generate_content.side_effect = api_error(status)
                self.assert_safe_error(expected)
                self.client.models.generate_content.assert_called_once()
                self.client.close.assert_called_once()
        self.sleep.assert_not_called()

    def test_invalid_key_http_400_is_auth(self):
        self.client.models.generate_content.side_effect = api_error(400, reason="API_KEY_INVALID")
        self.assert_safe_error("auth")
        self.client.close.assert_called_once()

    def test_sdk_malformed_response_is_sanitized(self):
        self.client.models.generate_content.side_effect = errors.UnknownApiResponseError(FAKE_KEY)
        self.assert_safe_error("invalid_response")
        self.client.close.assert_called_once()

    def test_client_initialization_failure_is_sanitized(self):
        self.factory.side_effect = RuntimeError(FAKE_KEY)
        self.assert_safe_error("request")
        self.client.close.assert_not_called()

    def test_timeout_is_sanitized_and_not_retried(self):
        self.client.models.generate_content.side_effect = TimeoutError(FAKE_KEY)
        self.assert_safe_error("timeout")
        self.client.models.generate_content.assert_called_once()
        self.client.close.assert_called_once()
        self.sleep.assert_not_called()

    def test_sdk_deadline_is_timeout_without_retry(self):
        self.client.models.generate_content.side_effect = api_error(504)
        self.assert_safe_error("timeout")
        self.client.models.generate_content.assert_called_once()
        self.client.close.assert_called_once()

    def test_temporary_unavailable_is_retried_once(self):
        self.client.models.generate_content.side_effect = [api_error(503), response("Recovered")]
        self.assertEqual(ai_service.generate_text("Lesson", FAKE_KEY), "Recovered")
        self.assertEqual(self.client.models.generate_content.call_count, 2)
        self.sleep.assert_called_once_with(ai_service.RETRY_DELAY_SECONDS)
        self.client.close.assert_called_once()

    def test_persistent_unavailable_stops_after_two_attempts(self):
        self.client.models.generate_content.side_effect = api_error(503)
        self.assert_safe_error("unavailable")
        self.assertEqual(self.client.models.generate_content.call_count, 2)
        self.client.close.assert_called_once()

    def test_connection_failure_is_retried_once(self):
        self.client.models.generate_content.side_effect = [ConnectionError(FAKE_KEY), response()]
        self.assertEqual(ai_service.generate_text("Lesson", FAKE_KEY), "Generated lesson")
        self.assertEqual(self.client.models.generate_content.call_count, 2)
        self.client.close.assert_called_once()

    def test_blocked_empty_and_truncated_outputs_are_rejected(self):
        for output, expected in [
            (response("unsafe", block_reason="SAFETY"), "blocked"),
            (response("unsafe", finish_reason=types.FinishReason.SAFETY), "blocked"),
            (response(""), "empty"),
            (response(None), "empty"),
            (response("Partial text", finish_reason="MAX_TOKENS"), "invalid_response"),
        ]:
            with self.subTest(expected=expected, output=output):
                self.client.reset_mock()
                self.client.models.generate_content.return_value = output
                self.assert_safe_error(expected)
                self.client.close.assert_called_once()
                self.client.models.generate_content.assert_called_once()

    def test_malformed_json_is_rejected(self):
        self.client.models.generate_content.return_value = response("{incomplete")
        self.assert_safe_error("invalid_response", lambda: ai_service.generate_text("Lesson", FAKE_KEY, json_mode=True))
        self.client.close.assert_called_once()

    def test_unspecified_block_reason_is_not_a_block(self):
        self.client.models.generate_content.return_value = response(
            block_reason=types.BlockedReason.BLOCKED_REASON_UNSPECIFIED
        )
        self.assertEqual(ai_service.generate_text("Lesson", FAKE_KEY), "Generated lesson")

    def test_unexpected_provider_error_and_cleanup_are_safe(self):
        self.client.models.generate_content.side_effect = RuntimeError(FAKE_KEY)
        self.client.close.side_effect = RuntimeError(FAKE_KEY)
        self.assert_safe_error("request")
        self.client.close.assert_called_once()

    def test_cleanup_failure_does_not_mask_success(self):
        self.client.close.side_effect = RuntimeError(FAKE_KEY)
        self.assertEqual(ai_service.generate_text("Lesson", FAKE_KEY), "Generated lesson")

    def test_separate_requests_create_separate_clients(self):
        second = MagicMock()
        second.models.generate_content.return_value = response("Second")
        self.factory.side_effect = [self.client, second]
        ai_service.generate_text("First", FAKE_KEY)
        self.assertEqual(ai_service.generate_text("Second", "another-test-key"), "Second")
        self.assertEqual(self.factory.call_count, 2)
        self.assertEqual(self.factory.call_args_list[0].kwargs["api_key"], FAKE_KEY)
        self.assertEqual(self.factory.call_args_list[1].kwargs["api_key"], "another-test-key")
        self.client.close.assert_called_once()
        second.close.assert_called_once()


class SettingsTests(unittest.TestCase):
    def test_personal_key_overrides_shared_settings(self):
        result = ai_service.resolve_settings(
            {"GEMINI_API_KEY": "env-key", "GEMINI_MODEL": "env-model"},
            {"DEFAULT_API_KEY": "shared-key", "GEMINI_MODEL": "secret-model"},
            user_key="  personal-key  ",
        )
        self.assertEqual(result, ("personal-key", "env-model", False))

    def test_existing_streamlit_secret_is_supported(self):
        self.assertEqual(
            ai_service.resolve_settings({}, {"DEFAULT_API_KEY": "shared-key"}),
            ("shared-key", ai_service.DEFAULT_MODEL, True),
        )

    def test_precedence_and_blank_values(self):
        self.assertEqual(
            ai_service.resolve_settings(
                {"GEMINI_API_KEY": " ", "GOOGLE_API_KEY": "env-key"},
                {"DEFAULT_API_KEY": "secret-key", "GEMINI_MODEL": "custom-model"},
            ),
            ("env-key", "custom-model", True),
        )
        self.assertEqual(ai_service.resolve_settings({}, {}), ("", ai_service.DEFAULT_MODEL, False))


if __name__ == "__main__":
    unittest.main()
