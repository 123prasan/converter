import sys
import os
import subprocess
import time
import threading

# ======================================================================================
#  HYPER-SPEED WORD-TO-PDF ENGINE v8.0 (The "Zero-Latency" Edition)
#  Optimizations:
#   1. Lazy Importing: Heavy modules load only when needed (Startup Boost)
#   2. Word COM Tuning: Disables Spellcheck/Pagination for raw speed
#   3. Micro-Polling: 10ms reaction time
#   4. Fast OS Checks: Removed 'platform' module overhead
# ======================================================================================

# --- CONFIGURATION ---
TIMEOUT_SECONDS = 25
POLL_INTERVAL = 0.01  # 10ms (Ultra-fast checks)

class SystemKernel:
    """Low-level OS manipulations for maximum performance"""
    
    @staticmethod
    def boost_process_priority():
        """Force High Priority to steal CPU cycles"""
        try:
            if os.name == 'nt': # Fast Windows check
                import ctypes
                # HIGH_PRIORITY_CLASS (0x00000080)
                ctypes.windll.kernel32.SetPriorityClass(
                    ctypes.windll.kernel32.GetCurrentProcess(), 0x00000080)
            else:
                os.nice(-10) # Linux/Mac High Priority
        except:
            pass 

    @staticmethod
    def get_fast_libreoffice_cmd(input_path, output_dir):
        """Generates the optimized command with RAM-disk profile"""
        # Lazy imports for speed
        import tempfile
        import uuid
        
        profile_dir = os.path.join(tempfile.gettempdir(), f"LO_{uuid.uuid4().hex[:8]}")
        
        if os.name == 'nt':
            from urllib.request import pathname2url
            uri = f"file:///{pathname2url(profile_dir)}"
            exe = 'soffice'
        else:
            uri = f"file://{profile_dir}"
            exe = 'libreoffice'

        cmd = [
            exe,
            f'-env:UserInstallation={uri}', 
            '--headless',
            '--nologo',       
            '--nodefault',    
            '--norestore',    
            '--writer',
            '--convert-to', 'pdf',
            '--outdir', output_dir,
            input_path
        ]
        return cmd, profile_dir

class ConversionRacer(threading.Thread):
    def __init__(self, name, target_func, args, success_event, result_holder):
        super().__init__(daemon=True)
        self.name = name
        self.target_func = target_func
        self.args = args
        self.success_event = success_event
        self.result_holder = result_holder
        self.process = None 

    def run(self):
        try:
            self.target_func(self)
        except Exception:
            pass

class EngineLogic:
    @staticmethod
    def run_subprocess(racer_obj):
        cmd = racer_obj.args['cmd']
        output_path = racer_obj.args['output_path']
        
        startupinfo = None
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
        racer_obj.process = subprocess.Popen(
            cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            startupinfo=startupinfo
        )
        
        # Micro-polling loop
        while racer_obj.process.poll() is None:
            if racer_obj.success_event.is_set():
                racer_obj.process.terminate()
                return
            time.sleep(POLL_INTERVAL)
            
        # Post-process check
        if os.path.exists(output_path) and os.path.getsize(output_path) > 100:
            if not racer_obj.success_event.is_set():
                racer_obj.result_holder['winner'] = racer_obj.name
                racer_obj.success_event.set()

    @staticmethod
    def run_com_automation(racer_obj):
        if os.name != 'nt': return
        
        input_path = racer_obj.args['input_path']
        output_path = racer_obj.args['output_path']
        
        # Lazy import COM libraries only inside the thread
        import pythoncom
        import win32com.client
        
        pythoncom.CoInitialize()
        
        try:
            try:
                word = win32com.client.GetActiveObject("Word.Application")
            except:
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                word.DisplayAlerts = False
                # SPEED TWEAK: Disable background tasks
                word.Options.CheckSpellingAsYouType = False
                word.Options.CheckGrammarAsYouType = False
                word.Options.Pagination = False
                word.ScreenUpdating = False
            
            # Open Read-Only for speed
            doc = word.Documents.Open(input_path, ReadOnly=True, Visible=False)
            doc.SaveAs(output_path, FileFormat=17) # 17 = PDF
            doc.Close(SaveChanges=False)
            
            if os.path.exists(output_path) and os.path.getsize(output_path) > 100:
                if not racer_obj.success_event.is_set():
                    racer_obj.result_holder['winner'] = "Word-COM"
                    racer_obj.success_event.set()
        except:
            pass
        finally:
            pythoncom.CoUninitialize()

def main():
    # 1. Start Clock & Boost Priority immediately
    start_time = time.time()
    SystemKernel.boost_process_priority()
    
    if len(sys.argv) < 3:
        sys.exit(1)
        
    input_path = os.path.abspath(sys.argv[1])
    output_path = os.path.abspath(sys.argv[2])
    output_dir = os.path.dirname(output_path)

    # 2. Prepare Racers
    success_event = threading.Event()
    result = {'winner': None}
    threads = []
    temp_dirs_to_clean = []

    # RACER 1: Windows Native COM (Optimized)
    if os.name == 'nt':
        t = ConversionRacer(
            "Word-COM", 
            EngineLogic.run_com_automation, 
            {'input_path': input_path, 'output_path': output_path},
            success_event, result
        )
        threads.append(t)

    # RACER 2: LibreOffice (Profile Optimized)
    lo_cmd, lo_profile = SystemKernel.get_fast_libreoffice_cmd(input_path, output_dir)
    temp_dirs_to_clean.append(lo_profile)
    
    t = ConversionRacer(
        "LibreOffice",
        EngineLogic.run_subprocess,
        {'cmd': lo_cmd, 'output_path': output_path}, 
        success_event, result
    )
    threads.append(t)

    # RACER 3: Pandoc (Fallback)
    t = ConversionRacer(
        "Pandoc",
        EngineLogic.run_subprocess,
        {'cmd': ['pandoc', input_path, '-o', output_path, '--pdf-engine=wkhtmltopdf'], 'output_path': output_path},
        success_event, result
    )
    threads.append(t)

    # 3. Start Race
    print(f"Status: Racing {len(threads)} engines...")
    sys.stdout.flush()
    
    for t in threads: t.start()

    # 4. Wait for Winner
    final_duration = 0
    lo_default_name = os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
    lo_default_path = os.path.join(output_dir, lo_default_name)

    while not success_event.is_set():
        if time.time() - start_time > TIMEOUT_SECONDS:
            break
        
        # Check for LibreOffice misnamed file (Common LO quirk)
        if os.path.exists(lo_default_path):
            # Double check size to ensure write is done
            if os.path.getsize(lo_default_path) > 100:
                try:
                    if os.path.exists(output_path): os.remove(output_path)
                    os.rename(lo_default_path, output_path)
                    result['winner'] = "LibreOffice-Detected"
                    success_event.set()
                except: pass
            
        time.sleep(POLL_INTERVAL)

    # 5. Capture Finish Time
    if success_event.is_set():
        final_duration = time.time() - start_time

    # 6. Aggressive Cleanup
    # Kill threads first
    for t in threads:
        if t.process:
            try: t.process.kill() 
            except: pass
    
    # Lazy import shutil only for cleanup
    import shutil
    if os.name == 'nt': time.sleep(0.1)
    for path in temp_dirs_to_clean:
        try: shutil.rmtree(path, ignore_errors=True)
        except: pass

    # 7. Final Output
    if success_event.is_set() and os.path.exists(output_path):
        print(f"Success: {result['winner']} won in {final_duration:.4f}s")
        sys.exit(0)
    else:
        print("Error: Conversion Timed Out")
        sys.exit(1)

if __name__ == "__main__":
    main()