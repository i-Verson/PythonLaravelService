import win32serviceutil
import win32service
import win32event
import servicemanager
import subprocess
import os
import sys
import threading
import time

class LaravelService(win32serviceutil.ServiceFramework):
    _svc_name_ = "DSCPayslip"
    _svc_display_name_ = "DscPayslipRun"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.processes = []
        self.stop_scheduler = False
        
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

    def run_scheduler(self):
        while not self.stop_scheduler:
            subprocess.run([self.php_path, "artisan", "schedule:run"], cwd=self.working_dir)
            for _ in range(60):
                if self.stop_scheduler:
                    break
                time.sleep(1)

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, "")
        )

        # Start Laravel serve
        self.processes.append(self.start_process("serve"))

        # Start schedule:run in a background thread
        self.stop_scheduler = False
        self.scheduler_thread = threading.Thread(target=self.run_scheduler)
        self.scheduler_thread.daemon = True
        self.scheduler_thread.start()

        # Start queue:work with optimizations
        self.processes.append(self.start_process("queue:work --sleep=10 --max-jobs=100 --max-time=3600"))

        # Wait until service is stopped
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.stop_scheduler = True
        if hasattr(self, 'scheduler_thread'):
            self.scheduler_thread.join(timeout=5)
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