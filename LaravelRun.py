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
        
        # Get the directory of the current executable
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            self.working_dir = os.path.dirname(sys.executable)
        else:
            # Running as script
            self.working_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Get PHP path using 'where php' command
        result = subprocess.run(['where', 'php'], capture_output=True, text=True, shell=True)
        if result.returncode == 0:
            self.php_path = result.stdout.strip().split('\n')[0]
        else:
            raise Exception("PHP not found in system PATH. Please ensure PHP is installed and added to PATH.")
        
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