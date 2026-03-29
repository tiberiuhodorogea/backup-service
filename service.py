"""
Windows Service wrapper for the Backup Service.

Install:   python service.py install
Start:     python service.py start
Stop:      python service.py stop
Remove:    python service.py remove

The service runs the FastAPI/uvicorn server on 0.0.0.0:8550.
"""

import os
import sys
import socket

# Ensure the project directory is on sys.path so imports work when
# running as a Windows service (working dir may differ).
SERVICE_DIR = os.path.dirname(os.path.abspath(__file__))
if SERVICE_DIR not in sys.path:
    sys.path.insert(0, SERVICE_DIR)

import win32serviceutil  # noqa: E402
import win32service      # noqa: E402
import win32event        # noqa: E402
import servicemanager    # noqa: E402


class BackupService(win32serviceutil.ServiceFramework):
    _svc_name_ = "BackupService"
    _svc_display_name_ = "Backup Service"
    _svc_description_ = (
        "Automatic file backup service with web dashboard. "
        "Access at http://<hostname>:8550"
    )

    def __init__(self, args):
        super().__init__(args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.server = None

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        if self.server:
            self.server.should_exit = True

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, ""),
        )
        self.main()

    def main(self):
        os.chdir(SERVICE_DIR)

        import asyncio
        import contextlib
        import uvicorn
        from main import app  # noqa: F811

        # Windows Service SvcDoRun() runs in a non-main thread, which hits two
        # restrictions:
        # 1. ProactorEventLoop.__init__ calls set_wakeup_fd() — main thread only.
        #    Fix: use SelectorEventLoop policy instead.
        # 2. uvicorn.Server.capture_signals calls signal.signal() — main thread only.
        #    Fix: replace capture_signals with a no-op context manager on the instance.
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

        config = uvicorn.Config(app, host="0.0.0.0", port=8550, log_level="info", log_config=None)
        self.server = uvicorn.Server(config)
        self.server.capture_signals = contextlib.nullcontext
        self.server.run()


if __name__ == "__main__":
    if len(sys.argv) == 1:
        # Called without args — SCM is starting the service
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(BackupService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(BackupService)
