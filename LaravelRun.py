import win32serviceutil
import win32service
import win32event
import servicemanager
import subprocess
import os
import sys
import shutil

class LaravelService(win32serviceutil.ServiceFramework):
    _svc_name_ = "DSCPayslip"
    _svc_display_name_ = "DscPayslipRun"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.processes = []

        # Dynamically get the path where the EXE or script is running
        self.working_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))

        # Attempt to detect PHP from system PATH
        self.php_path = shutil.which("php")
        if not self.php_path:
            raise FileNotFoundError("PHP executable not found in system PATH.")

    def start_process(self, command):
        return subprocess.Popen(
            [self.php_path, "artisan"] + command.split(),
            cwd=self.working_dir,
            stdout=subprocess.DEVNULL,  # Suppress console output (optional)
            stderr=subprocess.DEVNULL
        )

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, "")
        )

        servicemanager.LogInfoMsg(f"[{self._svc_name_}] Starting Laravel background workers...")
        servicemanager.LogInfoMsg(f"[{self._svc_name_}] Using PHP: {self.php_path}")
        servicemanager.LogInfoMsg(f"[{self._svc_name_}] Working directory: {self.working_dir}")

        # Start Laravel commands
        self.processes.append(self.start_process("serve"))
        self.processes.append(self.start_process("schedule:work"))
        self.processes.append(self.start_process("queue:work"))

        # Wait until service is stopped
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        servicemanager.LogInfoMsg(f"[{self._svc_name_}] Stopping Laravel workers...")

        for p in self.processes:
            p.terminate()

        self.ReportServiceStatus(win32service.SERVICE_STOPPED)
        win32event.SetEvent(self.hWaitStop)

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(LaravelService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(LaravelService)
