"""COM execution and Excel application management (Thread-local STA)."""
from __future__ import annotations

import threading
import pythoncom
import win32com.client as win32

_local = threading.local()


def get_excel_app(visible: bool = False):
    if not getattr(_local, "initialized", False):
        pythoncom.CoInitialize()
        _local.initialized = True

    app = getattr(_local, "app", None)
    if app is None:
        try:
            app = win32.DispatchEx("Excel.Application")
        except Exception:
            app = win32.Dispatch("Excel.Application")
        app.Visible = visible
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = 3  # Macros disabled by default
        _local.app = app
    return app


class DummyExecutor:
    def get_app(self, visible: bool = False):
        return get_excel_app(visible)
    def call(self, fn, *args, **kwargs):
        return run_com(fn, *args, **kwargs)
    def shutdown(self):
        shutdown_com()


_dummy = DummyExecutor()


def get_executor():
    return _dummy


def run_com(fn, *args, **kwargs):
    if not getattr(_local, "initialized", False):
        pythoncom.CoInitialize()
        _local.initialized = True
    return fn(*args, **kwargs)


def shutdown_com():
    app = getattr(_local, "app", None)
    if app is not None:
        try:
            app.Quit()
        except Exception:
            pass
        _local.app = None
    if getattr(_local, "initialized", False):
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        _local.initialized = False
