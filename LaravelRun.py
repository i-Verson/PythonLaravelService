import win32serviceutil
import win32service
import win32event
import servicemanager
import subprocess
import os
import sys

class LaravelService(win32serviceutil.ServiceFramework):
    _svc_name_ = "DSCPayslip"
    _svc_display_name_ = "DscPayslipRun"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.processes = []
        self.working_dir = r"C:\Users\rayiv\Documents\DSC Payroll Email Sender\DSCPayrollEmailSender"
        self.php_path = r"C:\php 8.3.22\php.exe"

    def start_process(self, command):
        return subprocess.Popen([self.php_path, "artisan"] + command.split(), cwd=self.working_dir)

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, "")
        )

        # Start Laravel commands
        self.processes.append(self.start_process("serve"))
        self.processes.append(self.start_process("schedule:work"))
        self.processes.append(self.start_process("queue:work"))

        # Wait until service is stopped
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
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