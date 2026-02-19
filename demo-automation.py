import time
import os
import json
import ctypes
from pywinauto import Application, Desktop

# --- CONFIGURATION ---
SHORTCUT_PATH = r"C:\Users\ASUS\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\MaxiFitPointV5.05.lnk"
CONFIG_PATH = r"d:\VS_CODE\Infiswift\MaxifitConfig.json"

def load_config():
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

CONFIG = load_config()

# --- UIA CACHE FOR SPEED ---
class UIACache:
    """Caches UIA connection to avoid reconnection overhead"""
    def __init__(self):
        self._uia_app = None
        self._hwnd = None
    
    def get_listbox(self, win):
        """Get ListBox via cached UIA connection"""
        hwnd = win.handle
        
        # Reuse connection if same window
        if self._hwnd != hwnd or self._uia_app is None:
            self._hwnd = hwnd
            self._uia_app = Application(backend="uia").connect(handle=hwnd)
        
        return self._uia_app.window(handle=hwnd).child_window(auto_id="pcsListBox")

# Global cache instance
_uia_cache = UIACache()

# --- WIN32 CONSTANTS ---
LB_GETCOUNT = 0x018B
LB_SETCURSEL = 0x0186

# --- HELPER FUNCTIONS ---

def connect_app():
    """Connect to existing MAXIFIT window"""
    print("Connecting to MAXIFIT...")
    try:
        app = Application(backend="win32").connect(title_re="MAXIFIT simulator", timeout=3)
        win = app.window(title_re="MAXIFIT simulator")
        if win.exists():
             print("  Connected Successfully.")
             return win
    except: pass
    return None

def click_btn(parent, auto_id):
    """Robustly clicks a button by auto_id"""
    print(f"  Clicking {auto_id}...")
    try:
        btn = parent.child_window(auto_id=auto_id)
        if not btn.exists():
             matches = parent.descendants(auto_id=auto_id)
             if matches: btn = matches[0]
             else:
                print(f"    Button {auto_id} not found.")
                return False
        
        # Click -> Invoke -> Physical Click sequence
        try: btn.click(); return True
        except: 
            try: btn.invoke(); return True
            except: 
                try: 
                    btn.set_focus()
                    btn.click_input()
                    return True
                except: pass
    except: pass
    return False

def select_pcs_item(main_win, index):
    """Select PCS list item using cached UIA connection"""
    print(f"  Selecting PCS Item at index {index} (Cached UIA)...")
    try:
        list_box = _uia_cache.get_listbox(main_win)
        items = list_box.children(control_type="ListItem")
        count = len(items)
        
        if count == 0:
             print("    List is empty.")
             return False
             
        # Resolve negative index
        target_idx = index if index >= 0 else count + index
        
        if target_idx < 0 or target_idx >= count:
            print(f"    Index {index} out of bounds (Count {count})")
            return False
            
        target_item = items[target_idx]
        print(f"    Found: {target_item.window_text()}")

        try:
            target_item.click_input()
        except:
            print("    click_input failed. Trying invoke...")
            target_item.invoke()
            
        time.sleep(0.5)
        return True

    except Exception as e:
        print(f"    Selection Error: {e}")
    return False

def clear_pcs_list(main_win):
    """Delete all items in PCS list using fast Win32 API"""
    print("  Clearing PCS List...")
    try:
        list_box = main_win.child_window(auto_id="pcsListBox")
        if list_box.exists():
            hwnd = list_box.handle
            for _ in range(10):
                # Fast Win32 count check
                count = ctypes.windll.user32.SendMessageW(hwnd, LB_GETCOUNT, 0, 0)
                if count <= 0:
                    break
                
                # Fast Win32 selection (select first item)
                ctypes.windll.user32.SendMessageW(hwnd, LB_SETCURSEL, 0, 0)
                time.sleep(0.1)
                
                click_btn(main_win, "pcsDeleteButton")
                time.sleep(0.3)
    except: pass

def set_combo(parent, auto_id, value, is_numeric=False, commit_key="{ENTER}"):
    """Set combo box or numeric up-down value"""
    print(f"  Setting {auto_id} -> {value}")
    try:
        ctrl = parent.child_window(auto_id=auto_id)
        if not ctrl.exists():
            print(f"    Error: Control {auto_id} not found.")
            return False
            
        ctrl.set_focus()
        time.sleep(0.3)

        if is_numeric:
            # Try ValuePattern then Key Typing
            try:
                 ctrl.wrapper_object().iface_value.SetValue(str(value))
                 time.sleep(0.1)
                 #if commit_key: ctrl.type_keys(commit_key)
                 return True
            except:
                #ctrl.type_keys("^a" + str(value) + commit_key, with_spaces=True) -->use of enter key is not needed
                ctrl.type_keys("^a" + str(value) , with_spaces=True)
                return True
        else:
            # Try Select then Expand/Click
            try:
                ctrl.select(str(value))
                return True
            except:
                try:
                    ctrl.expand(); time.sleep(0.5)
                    Desktop(backend="win32").window(title=str(value)).click_input()
                    return True
                except: pass
        return True
    except: return False

def handle_net_error():
    """Detect and dismiss .NET Framework error dialogs"""
    try:
        s = Desktop(backend="win32").windows(title_re=".*NET Framework.*|.*Exception.*")
        if s:
            error_dialog = s[0]
            print(f"    Found .NET Error: {error_dialog.window_text()}")
            error_dialog.set_focus()
            
            # Click Continue or Quit
            for title in ["Continue", "続行", "Quit", "終了"]:
                try:
                    btn = error_dialog.child_window(title=title, control_type="Button")
                    if btn.exists():
                        btn.click()
                        return True
                except: pass
            error_dialog.close()
            return True
    except: pass
    return False

def get_safe_filename(base_filename, check_extension=""):
    """
    Get a safe filename based on overwrite policy.
    
    Args:
        base_filename: The base name to save as (without extension)
        check_extension: The extension to check on disk (e.g., '.csv', '.pdf', '.sav')
    
    Returns:
        Safe filename (without extension) for the save dialog
    
    If overwrite_existing is False and file exists on disk, auto-increment (file_1, file_2, etc.)
    """
    overwrite = CONFIG.get('output_files', {}).get('overwrite_existing', False)
    output_dir = CONFIG.get('output_files', {}).get('output_directory', '')
    
    if overwrite:
        return base_filename
    
    # Build full path for checking
    if output_dir:
        disk_filename = os.path.join(output_dir, base_filename + check_extension) if check_extension else os.path.join(output_dir, base_filename)
    else:
        disk_filename = base_filename + check_extension if check_extension else base_filename
    
    if not os.path.exists(disk_filename):
        return base_filename
    
    # File exists, auto-increment
    counter = 1
    while True:
        new_base = f"{base_filename}_{counter}"
        
        if output_dir:
            new_disk = os.path.join(output_dir, new_base + check_extension) if check_extension else os.path.join(output_dir, new_base)
        else:
            new_disk = new_base + check_extension if check_extension else new_base
        
        if not os.path.exists(new_disk):
            print(f"    File exists. Using: {new_base}")
            return new_base
        
        counter += 1
        if counter > 100:  # Safety limit
            return base_filename

def handle_save_dialog(filename):
    """Handle 'Save As' and confirmation dialogs"""
    print(f"  Handling Save Dialog for '{filename}'...")
    savewin = None
    
    for _ in range(20): # 10s timeout
        time.sleep(0.5)
        try:
            s = Desktop(backend="win32").windows(title_re=".*名前を付けて保存.*|.*Save As.*|.*保存.*|.*Save Print.*|.*確認.*")
            if s:
                savewin = s[0]
                # If it's just a confirmation, click Yes and keep looking
                if "Confirm" in savewin.window_text() or "確認" in savewin.window_text():
                    print("    Confirm Dialog Found. Clicking Yes.")
                    for title in ["Yes", "はい(Y)"]:
                        try:
                            btn = savewin.child_window(title=title, control_type="Button")
                            if btn.exists(): btn.click(); break
                        except: pass
                    continue
                break
        except: pass

    if savewin:
        print(f"    Dialog: {savewin.window_text()}")
        try:
            savewin.set_focus()
            handle_net_error() # Clear any blocking errors
            
            # Navigate to output directory if configured
            output_dir = CONFIG.get('output_files', {}).get('output_directory', '')
            if output_dir:
                print(f"    Navigating to: {output_dir}")
                # Ctrl+L focuses address bar in save dialog
                try:
                    savewin.type_keys("^l", with_spaces=True)
                    time.sleep(0.3)
                    savewin.type_keys(output_dir, with_spaces=True)
                    time.sleep(0.3)
                    savewin.type_keys("{ENTER}")
                    time.sleep(0.5)
                except Exception as e:
                    print(f"    Directory navigation failed: {e}")
            
            # Alt+N to focus filename
            try: savewin.type_keys("%n", with_spaces=True)
            except: pass
            time.sleep(0.3)
            
            # Type filename and Enter
            savewin.type_keys(filename, with_spaces=True)
            time.sleep(0.3)
            savewin.type_keys("{ENTER}")
            
            # Catch immediate Overwrite prompt
            time.sleep(1.0)
            try:
                conf = Desktop(backend="win32").window(title_re=".*Save Print.*|.*確認.*")
                if conf.exists(timeout=2):
                    conf.set_focus()
                    conf.type_keys("{ENTER}") # Accept Overwrite
            except: pass
            return True
        except Exception as e:
            print(f"    Save Error: {e}")
    else:
        print("    Save dialog did not appear.")
    return False

# --- MAIN WORKFLOW ---

def main():
    print("--- Starting Automation ---")
    
    # Try to connect to existing window first
    main_win = connect_app()
    
    if not main_win:
        # Launch if not already open
        print("Launching MaxiFit...")
        if not os.path.exists(SHORTCUT_PATH):
            print("Shortcut not found!")
            return
            
        os.startfile(SHORTCUT_PATH)
        
        # Wait for window to appear
        for i in range(15):
            print(f"  Waiting for UI... {i+1}/15")
            time.sleep(1)
            main_win = connect_app()
            if main_win: break
    else:
        print("Connected to existing window.")
    
    if not main_win:
        print("Failed to connect to application.")
        return

    main_win.set_focus()
    time.sleep(1)
    clear_pcs_list(main_win)

    # STAGE 1: Location
    print("\n[1] Location")
    if not set_combo(main_win, "areaSelectComboBox", CONFIG["prefecture"]): return
    time.sleep(0.5)
    if not set_combo(main_win, "pointSelectComboBox", CONFIG["subregion"]): return
    time.sleep(0.5)
    
    # STAGE 2: Add First PCS
    print("\n[2] Adding First PCS")
    first_array = CONFIG['pv_arrays'][0]
    print(f"  PCS: {first_array['pcs']}")
    
    if not set_combo(main_win, "pcsSelectComboBox", first_array['pcs']): return
    click_btn(main_win, "pcsAddButton")
    time.sleep(0.5)
    select_pcs_item(main_win, 0)
    time.sleep(0.5)

    # STAGE 3: PV Array Configuration
    print("\n[3] PV Array Setup")
    
    for array_idx, array_config in enumerate(CONFIG['pv_arrays']):
        print(f"\n  === Array Type {array_idx + 1} ===")
        print(f"    PCS: {array_config['pcs']}")
        print(f"    Panel: {array_config['panel_type']}")
        print(f"    Series: {array_config['panel_series']}, Parallel: {array_config['panel_parallel']}")
        print(f"    Angle: {array_config['placement_angle']}°, Direction: {array_config['direction']}°")
        print(f"    Total Arrays: {array_config['num_arrays']}")
        
        # For array types after the first, add new PCS
        if array_idx > 0:
            print(f"  Adding new PCS type...")
            
            if not set_combo(main_win, "pcsSelectComboBox", array_config['pcs']): 
                print("    Failed to select PCS")
                continue
            
            click_btn(main_win, "pcsAddButton")
            time.sleep(0.5)
            select_pcs_item(main_win, -1)
            time.sleep(0.5)
        
        # Configure panel settings
        print(f"  Configuring panel settings...")
        click_btn(main_win, "panelSettingButton01")
        time.sleep(1.5)
        
        popup = Desktop(backend="win32").window(title_re=".*パネル入力.*")
        if popup.exists(timeout=5):
            popup.set_focus()
            set_combo(popup, "panelSelectComboBoxSub", array_config['panel_type'])
            set_combo(popup, "panelSeriesNumericUpDown", array_config['panel_series'], True)
            set_combo(popup, "panelParallelNumericUpDown", array_config['panel_parallel'], True)
            set_combo(popup, "installationAngleNumericUpDown", array_config['placement_angle'], True)
            set_combo(popup, "azimuthNumericUpDown", array_config['direction'], True)
            #set_combo(popup, "backSideRateNumericUpDown", array_config['backside_efficiency'], True)
            
            click_btn(popup, "enterButton")
            time.sleep(1)
        
        # Copy this configured PCS
        num_copies = array_config['num_arrays'] - 1
        if num_copies > 0:
            print(f"  Creating {num_copies} copies...")
            for _ in range(num_copies):
                click_btn(main_win, "pcsCopyButton")
                time.sleep(0.3)

    # STAGE 4: Snowfall
    print("\n[4] Snowfall")
    try:
        chk = main_win.child_window(auto_id="snowFlagCheck")
        if chk.exists() and chk.get_toggle_state() == 1: chk.toggle()
    except: pass

    # STAGE 5: Efficiency
    print("\n[5] Efficiency")
    '''
    click_btn(main_win, "performSetButton")
    time.sleep(1.5)
    
    popup = Desktop(backend="win32").window(title_re=".*効率.*")
    if popup.exists(timeout=5):
        popup.set_focus()
        popup.type_keys(f"^a{CONFIG['system_efficiency']}{{TAB}}^a{CONFIG['power_efficiency']}{{TAB}}{{ENTER}}", with_spaces=True)
        time.sleep(1)
    '''
    # STAGE 6: Export & Excel
    print("\n[6] Export & Excel")
    #set_combo(main_win, "dataTypeComboBox", "平均年")
    
    # Retry loop for Chart button
    chart = None
    for _ in range(3):
        click_btn(main_win, "totalChartViewButton")
        time.sleep(1.5)
        try:
            cands = Desktop(backend="win32").windows(title_re=".*トータルチャート.*")
            if cands:
                chart = cands[0]
                print(f"  Chart Popup: {chart.window_text()}")
                break
        except: pass

    if chart:
        chart.set_focus()
        
        # CSV Export
        print("  Exporting CSV...")
        csv_filename = CONFIG.get('output_files', {}).get('csv_filename', 'output')
        safe_csv = get_safe_filename(csv_filename, '.csv')
        
        for child in chart.descendants():
            if "CSV書き出し" in child.window_text():
                child.click_input()
                handle_save_dialog(safe_csv)
                break
        
        # Excel Print
        print("  Triggering Excel Print...")
        #os.system("taskkill /f /im excel.exe >nul 2>&1")
        time.sleep(1)
        
        # Find and click Total Print button
        for child in chart.descendants():
            if "トータルプリント" in child.window_text():
                print("    Clicking [Total Print]...")
                child.click_input()
                break

        # Wait for Excel
        print("    Waiting for Excel...")
        excel = None
        for _ in range(30):
            print(".", end="", flush=True)
            time.sleep(1)
            try:
                s = Desktop(backend="uia").windows(title_re=".*Excel.*")
                if s: excel = s[0]; break
            except: pass
        print()
        
        if excel:
            print(f"    Excel Found: {excel.window_text()}")
            
            # Focus Excel with retry
            excel_ready = False
            for _ in range(10):
                try:
                    excel.set_focus()
                    time.sleep(1)
                    excel_ready = True
                    break
                except Exception as e:
                    print(f"    Excel Busy: {e}")
                    time.sleep(1)
            
            if excel_ready:
                # Search for Print button
                print("    Searching for Print Button...")
                print_btn = None
                for title in ["Print", "印刷", "Quick Print", "クイック印刷"]:
                    try:
                        matches = excel.descendants(title=title, control_type="Button")
                        if matches:
                            print_btn = matches[0]
                            print(f"    Found: {title}")
                            break
                    except: pass
                handle_net_error()
                if print_btn:
                    print_btn.click_input()
                else:
                    print("    Using Ctrl+P...")
                    excel.type_keys("^p", with_spaces=True)
                    time.sleep(2)
            else:
                print("    Excel not responding. Trying blind Ctrl+P...")
                try: excel.type_keys("^p", with_spaces=True); time.sleep(2)
                except: pass

            # Confirm Print
            try: excel.type_keys("{ENTER}")
            except: pass
            
            print_filename = CONFIG.get('output_files', {}).get('print_filename', 'print_output')
            safe_print = get_safe_filename(print_filename, '.pdf')
            handle_save_dialog(safe_print)
        else:
            print("    Excel did not appear.")

    # STAGE 7: Save Config
    print("\n[7] Save Config")
    main_win.set_focus()
    time.sleep(0.5)
    click_btn(main_win, "fileSaveButton")
    
    config_filename = CONFIG.get('output_files', {}).get('config_filename', 'MAXFITconfig')
    safe_config = get_safe_filename(config_filename, '.sav')
    handle_save_dialog(safe_config)

    print("\n--- Automation Complete ---")

if __name__ == "__main__":
    main()
