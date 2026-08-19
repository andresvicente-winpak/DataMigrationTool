from unittest.mock import patch

from modules.gui.app import M3MigrationApp


def test_user_guide_opens_default_mkdocs_url():
    with patch("modules.gui.app.webbrowser.open_new_tab", return_value=True) as open_tab:
        M3MigrationApp.open_documentation(object())

    open_tab.assert_called_once_with("http://localhost:8000/")


def test_user_guide_url_can_be_configured(monkeypatch):
    monkeypatch.setenv("M3_USER_GUIDE_URL", "https://docs.example.com/migration/")

    with patch("modules.gui.app.webbrowser.open_new_tab", return_value=True) as open_tab:
        M3MigrationApp.open_documentation(object())

    open_tab.assert_called_once_with("https://docs.example.com/migration/")


def test_user_guide_reports_browser_error():
    with (
        patch("modules.gui.app.webbrowser.open_new_tab", return_value=False),
        patch("modules.gui.app.messagebox.showerror") as show_error,
    ):
        M3MigrationApp.open_documentation(object())

    show_error.assert_called_once()
    assert "mkdocs serve" in show_error.call_args.args[1]
