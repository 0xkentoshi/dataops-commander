import asyncio
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch


os.environ.setdefault("TELEGRAM_BOT_TOKEN", "123456:TEST_TOKEN")

try:
    import aiogram  # noqa: F401
except (ModuleNotFoundError, ImportError):
    aiogram_mod = types.ModuleType("aiogram")

    class _FilterExpr:
        def __getattr__(self, name):
            return self
        def __eq__(self, other):
            return self
        def startswith(self, *args, **kwargs):
            return self

    class _Router:
        def message(self, *args, **kwargs):
            return lambda fn: fn
        def callback_query(self, *args, **kwargs):
            return lambda fn: fn

    class _Dispatcher:
        def include_router(self, *args, **kwargs):
            pass
        def resolve_used_update_types(self):
            return []

    class _Bot:
        pass

    aiogram_mod.Bot = _Bot
    aiogram_mod.Dispatcher = _Dispatcher
    aiogram_mod.F = _FilterExpr()
    aiogram_mod.Router = _Router
    sys.modules["aiogram"] = aiogram_mod

    client_mod = types.ModuleType("aiogram.client")
    default_mod = types.ModuleType("aiogram.client.default")
    class DefaultBotProperties:
        def __init__(self, *args, **kwargs):
            pass
    default_mod.DefaultBotProperties = DefaultBotProperties
    sys.modules["aiogram.client"] = client_mod
    sys.modules["aiogram.client.default"] = default_mod

    enums_mod = types.ModuleType("aiogram.enums")
    class ParseMode:
        HTML = "HTML"
    enums_mod.ParseMode = ParseMode
    sys.modules["aiogram.enums"] = enums_mod

    filters_mod = types.ModuleType("aiogram.filters")
    class Command:
        def __init__(self, *args, **kwargs):
            pass
    class CommandStart:
        def __init__(self, *args, **kwargs):
            pass
    filters_mod.Command = Command
    filters_mod.CommandStart = CommandStart
    sys.modules["aiogram.filters"] = filters_mod

    types_mod = types.ModuleType("aiogram.types")
    class _SimpleType:
        def __init__(self, *args, **kwargs):
            for key, value in kwargs.items():
                setattr(self, key, value)
    for name in [
        "CallbackQuery", "FSInputFile", "InlineKeyboardButton",
        "InlineKeyboardMarkup", "Message"
    ]:
        setattr(types_mod, name, type(name, (_SimpleType,), {}))
    sys.modules["aiogram.types"] = types_mod

try:
    import ollama  # noqa: F401
except (ModuleNotFoundError, ImportError):
    ollama_stub = types.ModuleType("ollama")

    class AsyncClientStub:
        def __init__(self, *args, **kwargs) -> None:
            pass

    ollama_stub.AsyncClient = AsyncClientStub
    sys.modules["ollama"] = ollama_stub

from app import main as app_main
from app.database import DataRepository
from app.source_catalog import SourceCatalog


class FakeBot:
    def __init__(self, next_message_id: int = 900) -> None:
        self.next_message_id = next_message_id
        self.edits: list[dict] = []
        self.sent: list[dict] = []
        self.pinned: list[dict] = []
        self.unpinned: list[dict] = []

    async def edit_message_text(self, **kwargs):
        self.edits.append(kwargs)

    async def send_message(self, **kwargs):
        self.sent.append(kwargs)
        return SimpleNamespace(message_id=self.next_message_id)

    async def pin_chat_message(self, **kwargs):
        self.pinned.append(kwargs)

    async def unpin_chat_message(self, **kwargs):
        self.unpinned.append(kwargs)


class FakeCallback:
    def __init__(self, message_id: int) -> None:
        self.data = "files:page:0"
        self.from_user = SimpleNamespace(id=42)
        self.message = SimpleNamespace(
            message_id=message_id,
            chat=SimpleNamespace(id=42),
        )
        self.answers: list[tuple] = []

    async def answer(self, *args, **kwargs) -> None:
        self.answers.append((args, kwargs))


class FakeCommandMessage:
    def __init__(self, user_id: int = 42) -> None:
        self.from_user = SimpleNamespace(id=user_id)
        self.chat = SimpleNamespace(id=user_id)
        self.answers: list[tuple[tuple, dict]] = []

    async def answer(self, *args, **kwargs):
        self.answers.append((args, kwargs))
        return SimpleNamespace(message_id=901)


class DashboardNavigationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.catalog = SourceCatalog(
            self.root / "workspace",
            self.root / "telegram",
        )
        self.catalog.initialize()
        self.repository = DataRepository(self.root / "audit.sqlite3")
        self.repository.initialize()
        self.repository.accept_onboarding(
            42, app_main.CURRENT_ONBOARDING_VERSION
        )
        self.repository.set_dashboard(42, 42, 111, 0)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_back_button_opens_list_in_clicked_message(self) -> None:
        bot = FakeBot()
        callback = FakeCallback(message_id=222)

        with (
            patch.object(app_main, "catalog", self.catalog),
            patch.object(app_main, "repository", self.repository),
            patch.object(app_main, "allowed", return_value=True),
        ):
            asyncio.run(app_main.files_callback(callback, bot))

        self.assertEqual([item["message_id"] for item in bot.edits], [222])
        self.assertEqual(bot.sent, [])
        self.assertEqual(bot.unpinned[0]["message_id"], 111)
        self.assertEqual(bot.pinned[-1]["message_id"], 222)
        self.assertEqual(self.repository.get_dashboard(42).message_id, 222)


    def test_files_command_is_blocked_before_onboarding(self) -> None:
        bot = FakeBot(next_message_id=333)
        message = FakeCommandMessage(user_id=43)

        with (
            patch.object(app_main, "catalog", self.catalog),
            patch.object(app_main, "repository", self.repository),
            patch.object(app_main, "allowed", return_value=True),
        ):
            asyncio.run(app_main.files(message, bot))

        self.assertEqual(bot.sent, [])
        self.assertEqual(bot.edits, [])
        self.assertEqual(len(message.answers), 1)
        self.assertIn("/start", message.answers[0][0][0])
        self.assertEqual(self.repository.get_onboarding_version(43), 0)

    def test_files_command_creates_fresh_visible_dashboard(self) -> None:
        bot = FakeBot(next_message_id=333)
        message = FakeCommandMessage()

        with (
            patch.object(app_main, "catalog", self.catalog),
            patch.object(app_main, "repository", self.repository),
            patch.object(app_main, "allowed", return_value=True),
        ):
            asyncio.run(app_main.files(message, bot))

        self.assertEqual(bot.edits, [])
        self.assertEqual(len(bot.sent), 1)
        self.assertEqual(bot.unpinned[0]["message_id"], 111)
        self.assertEqual(bot.pinned[-1]["message_id"], 333)
        self.assertEqual(self.repository.get_dashboard(42).message_id, 333)


if __name__ == "__main__":
    unittest.main(verbosity=2)
