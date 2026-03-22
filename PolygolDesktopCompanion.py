VERSION = "1.0.1"

import asyncio
import websockets
import json
import pyautogui
import pyperclip
import platform
import subprocess
import base64
import socket
import os
import sys
import shutil
import threading

try:
    import pystray
    from PIL import Image
except ImportError:
    print("WARNING: 'pystray' or 'Pillow' not installed. System tray won't work.")
    print("Run: pip install pystray Pillow")

try:
    import tkinter as tk
    from tkinter import messagebox
except ImportError:
    tk = None

pyautogui.FAILSAFE = False
OS_TYPE = platform.system()

WIN_MEDIA_SUPPORTED = False
if OS_TYPE == "Windows":
    try:
        from winsdk.windows.media.control import GlobalSystemMediaTransportControlsSessionManager
        from winsdk.windows.storage.streams import DataReader, Buffer, InputStreamOptions
        WIN_MEDIA_SUPPORTED = True
    except ImportError:
        print("WARNING: 'winsdk' not installed. Media metadata won't work on Windows.")
        print("Run: pip install winsdk")

CURRENT_MEDIA_APP_ID = None

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and PyInstaller """
    try:
        base_path = sys._MEIPASS  # PyInstaller extracts files here
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

async def get_windows_media_info():
    global CURRENT_MEDIA_APP_ID
    if not WIN_MEDIA_SUPPORTED:
        return None
    try:
        manager = await GlobalSystemMediaTransportControlsSessionManager.request_async()
        
        session = manager.get_current_session()
        
        if not session:
            sessions = manager.get_sessions()
            if sessions and len(sessions) > 0:
                session = sessions[0]
        
        if not session:
            return None
        
        source_app = session.source_app_user_model_id
        CURRENT_MEDIA_APP_ID = source_app

        app_name_lower = source_app.lower()
        if "spotify" in app_name_lower: host_app = "spotify"
        elif "chrome" in app_name_lower: host_app = "chrome"
        elif "msedge" in app_name_lower: host_app = "edge"
        elif "firefox" in app_name_lower: host_app = "firefox"
        elif "music.ui" in app_name_lower or "itunes" in app_name_lower: host_app = "apple_music"
        elif "vlc" in app_name_lower: host_app = "vlc"
        else: host_app = "unknown"
        info = await session.try_get_media_properties_async()
        playback_info = session.get_playback_info()
        timeline = session.get_timeline_properties()

        duration = timeline.end_time.total_seconds()
        position = timeline.position.total_seconds()

        status = playback_info.playback_status
        state = "playing" if status == 4 else "paused"

        thumbnail_b64 = ""
        if info.thumbnail:
            try:
                ref = await info.thumbnail.open_read_async()
                buffer = Buffer(ref.size)
                await ref.read_async(buffer, buffer.capacity, 0)
                reader = DataReader.from_buffer(buffer)
                bytes_data = bytearray(ref.size)
                reader.read_bytes(bytes_data)
                thumbnail_b64 = "data:image/jpeg;base64," + base64.b64encode(bytes_data).decode('utf-8')
            except Exception as e:
                print(f"[Media] Thumbnail read error: {e}")

        return {
            "module": "media",
            "title": info.title or "Unknown Title",
            "artist": info.artist or "Unknown Artist",
            "album": info.album_title or "",
            "state": state,
            "duration": duration,
            "position": position,
            "app": source_app,
            "hostApp": host_app,
            "albumArt": thumbnail_b64
        }

    except Exception as e:
        print(f"[SysLink] Error fetching Windows Media: {e}")
        return None

async def get_macos_media_info():
    global CURRENT_MEDIA_APP_ID
    try:
        script = """
        if application "Spotify" is running then
            tell application "Spotify"
                set t to name of current track
                set a to artist of current track
                set s to player state
                return t & "|||" & a & "|||" & s & "|||Spotify"
            end tell
        else if application "Music" is running then
            tell application "Music"
                set t to name of current track
                set a to artist of current track
                set s to player state
                return t & "|||" & a & "|||" & s & "|||Music"
            end tell
        end if
        return ""
        """
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True).stdout.strip()
        if result:
            parts = result.split("|||")
            if len(parts) >= 4:
                state = "playing" if parts[2].lower() == "playing" else "paused"
                CURRENT_MEDIA_APP_ID = parts[3]
                host_app = "spotify" if parts[3] == "Spotify" else "apple_music"
                return {
                    "module": "media",
                    "title": parts[0],
                    "artist": parts[1],
                    "state": state,
                    "hostApp": host_app,
                    "albumArt": ""
                }
    except Exception:
        pass
    return None

async def broadcast_state(websocket):
    import psutil
    last_song_identity = ""
    print("[SysLink] Broadcast loop started.")
    
    await websocket.send(json.dumps({"module": "system", "action": "log", "message": "SysLink Core Handshake Complete"}))

    while True:
        try:
            media_info = None
            if OS_TYPE == "Windows":
                media_info = await get_windows_media_info()
            elif OS_TYPE == "Darwin":
                media_info = await get_macos_media_info()

            if media_info:
                current_identity = f"{media_info['title']}{media_info['artist']}"
                
                if current_identity != last_song_identity:
                    await websocket.send(json.dumps(media_info))
                    last_song_identity = current_identity
                else:
                    update_payload = media_info.copy()
                    update_payload["albumArt"] = None
                    await websocket.send(json.dumps(update_payload))
            else:
                if last_song_identity != "":
                    await websocket.send(json.dumps({"module": "media", "action": "clear"}))
                    last_song_identity = ""

            await asyncio.sleep(1.5)
        except Exception as e:
            print(f"[SysLink] Broadcast error: {e}")
            break

async def handle_message(websocket, message):
    try:
        payload = json.loads(message)
        module = payload.get("module")
        action = payload.get("action")
        
        if module == "input":
            screen_w, screen_h = pyautogui.size()
            if action == "move":
                pyautogui.moveTo(int(payload['x'] * screen_w), int(payload['y'] * screen_h))
            elif action == "down":
                pyautogui.mouseDown(button='right' if payload.get('button') == 2 else 'left')
            elif action == "up":
                pyautogui.mouseUp(button='right' if payload.get('button') == 2 else 'left')
            elif action == "scroll":
                pyautogui.scroll(int(-payload.get('dy', 0)))
            elif action == "keydown":
                key = payload.get('key')
                if key == "Enter": pyautogui.press('enter')
                elif key == "Backspace": pyautogui.press('backspace')
                elif len(key) == 1: pyautogui.keyDown(key.lower())
            elif action == "keyup":
                key = payload.get('key')
                if len(key) == 1: pyautogui.keyUp(key.lower())

        elif module == "media":
            if action == "playPause":
                if OS_TYPE == "Darwin": subprocess.run("osascript -e 'tell application \"System Events\" to key code 100'", shell=True)
                else: pyautogui.press('playpause')
            elif action == "next":
                if OS_TYPE == "Darwin": subprocess.run("osascript -e 'tell application \"System Events\" to key code 101'", shell=True)
                else: pyautogui.press('nexttrack')
            elif action == "prev":
                if OS_TYPE == "Darwin": subprocess.run("osascript -e 'tell application \"System Events\" to key code 98'", shell=True)
                else: pyautogui.press('prevtrack')
            elif action == "openApp":
                if OS_TYPE == "Windows" and CURRENT_MEDIA_APP_ID:
                    subprocess.run(["explorer", f"shell:AppsFolder\\{CURRENT_MEDIA_APP_ID}"])
                elif OS_TYPE == "Darwin" and CURRENT_MEDIA_APP_ID:
                    subprocess.run(["osascript", "-e", f'tell application "{CURRENT_MEDIA_APP_ID}" to activate'])

        elif module == "power":
            if action == "lock":
                if OS_TYPE == "Windows": subprocess.run("rundll32.exe user32.dll,LockWorkStation")
                elif OS_TYPE == "Darwin": subprocess.run("/System/Library/CoreServices/Menu Extras/User.menu/Contents/Resources/CGSession -suspend", shell=True)

        elif module == "hardware":
            if action == "volumeUp":
                pyautogui.press('volumeup')
            elif action == "volumeDown":
                pyautogui.press('volumedown')
            elif action == "brightnessUp":
                if OS_TYPE == "Windows":
                    subprocess.run('powershell "(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1, ((Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness + 10))"', shell=True)
                elif OS_TYPE == "Darwin":
                    subprocess.run("osascript -e 'tell application \"System Events\" to key code 144'", shell=True)
            elif action == "brightnessDown":
                if OS_TYPE == "Windows":
                    subprocess.run('powershell "(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1, ((Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness - 10))"', shell=True)
                elif OS_TYPE == "Darwin":
                    subprocess.run("osascript -e 'tell application \"System Events\" to key code 145'", shell=True)

        elif module == "run" or module == "shell" or module == "cmd":
            try:
                subprocess.Popen(action, shell=True)
            except Exception as e:
                print(f"[SysLink] Error running command: {e}")

        elif module == "clipboard":
            if action == "write":
                pyperclip.copy(payload.get("text", ""))
            elif action == "read":
                await websocket.send(json.dumps({"module": "clipboard", "action": "sync", "text": pyperclip.paste()}))

    except Exception as e:
        print(f"Error handling message: {e}")

# Global config cache
SYSLINK_CONFIG = {}
if os.path.exists(CONFIG_FILE):
    with open(CONFIG_FILE, 'r') as f:
        SYSLINK_CONFIG = json.load(f)

async def handler(websocket):
    global SYSLINK_CONFIG
    expected_token = SYSLINK_CONFIG.get("auth_token")
    
    try:
        # Increase timeout to 5 seconds for slower networks
        auth_payload_raw = await asyncio.wait_for(websocket.recv(), timeout=5.0)
        auth_payload = json.loads(auth_payload_raw)
        
        client_token = auth_payload.get("token")
        
        if auth_payload.get("module") == "auth" and client_token == expected_token:
            await websocket.send(json.dumps({"module": "auth", "status": "success", "version": VERSION}))
            print(f"[Auth] Success: {websocket.remote_address[0]}")
        else:
            print(f"[Auth] Failed: Invalid token from {websocket.remote_address[0]}")
            await websocket.send(json.dumps({"module": "auth", "status": "failed"}))
            await websocket.close()
            return
    except Exception as e:
        print(f"[Auth] Error during handshake: {e}")
        await websocket.close()
        return

    try:
        # Wait for the first message which MUST be the auth token
        # If no auth received within 3 seconds, drop connection
        auth_payload_raw = await asyncio.wait_for(websocket.recv(), timeout=3.0)
        auth_payload = json.loads(auth_payload_raw)
        
        if auth_payload.get("module") == "auth" and auth_payload.get("token") == expected_token:
            authenticated = True
            await websocket.send(json.dumps({"module": "auth", "status": "success", "version": VERSION}))
            print(f"--- Client Authenticated: {websocket.remote_address[0]} ---")
        else:
            await websocket.send(json.dumps({"module": "auth", "status": "failed"}))
            await websocket.close()
            return
    except Exception:
        await websocket.close()
        return

    stats_task = asyncio.create_task(broadcast_state(websocket))
    try:
        async for message in websocket:
            await handle_message(websocket, message)
    except websockets.exceptions.ConnectionClosed:
        print(f"Client disconnected: {websocket.remote_address[0]}")
    finally:
        stats_task.cancel()

# --- Configuration & Paths ---
if OS_TYPE == "Windows":
    APP_DIR = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'PolygolDesktopCompanion')
else:
    APP_DIR = os.path.expanduser('~/.polygol_companion')

CONFIG_FILE = os.path.join(APP_DIR, 'config.json')

def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('10.255.255.255', 1))
        IP = s.getsockname()[0]
    except Exception:
        IP = '127.0.0.1'
    finally:
        s.close()
    return IP

def add_to_startup(file_path):
    if OS_TYPE == "Windows":
        import winreg
        key = winreg.HKEY_CURRENT_USER
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            registry_key = winreg.OpenKey(key, key_path, 0, winreg.KEY_WRITE)
            winreg.SetValueEx(registry_key, "PolygolCompanion", 0, winreg.REG_SZ, file_path)
            winreg.CloseKey(registry_key)
        except Exception as e:
            print(f"Failed to add to startup: {e}")

# --- GUI Setup ---
def run_setup():
    if not tk:
        print("Tkinter not found. Skipping Setup GUI.")
        return

    root = tk.Tk()
    root.title("Polygol Desktop Companion Setup")
    root.geometry("450x300")
    root.eval('tk::PlaceWindow . center')
    root.configure(bg="#1c1c1c")

    tk.Label(root, text="Polygol Desktop Companion", font=("Arial", 16, "bold"), bg="#1c1c1c", fg="#ffffff").pack(pady=15)
    tk.Label(root, text=f"Local IP: {get_local_ip()}", font=("Arial", 12), bg="#1c1c1c", fg="#b3b3b3").pack(pady=5)
    tk.Label(root, text=f"Ver {VERSION}.", font=("Arial", 12), bg="#1c1c1c", fg="#b3b3b3").pack(pady=5)
    tk.Label(root, text=f"This product and Polygol is made by kirbIndustries", font=("Arial", 12), bg="#1c1c1c", fg="#b3b3b3").pack(pady=5)

    move_var = tk.BooleanVar(value=True)
    startup_var = tk.BooleanVar(value=True)

    if OS_TYPE == "Windows":
        tk.Checkbutton(root, text="Move to a permanent location (Recommended)", variable=move_var, bg="#1c1c1c", fg="#ffffff", selectcolor="#333333", activebackground="#1c1c1c").pack(anchor="w", padx=80)

    tk.Checkbutton(root, text="Run at Startup", variable=startup_var, bg="#1c1c1c", fg="#ffffff", selectcolor="#333333", activebackground="#1c1c1c").pack(anchor="w", padx=80)

    def on_save():
        import secrets
        os.makedirs(APP_DIR, exist_ok=True)
        # Generate a 32-character secure token
        token = secrets.token_hex(16)
        with open(CONFIG_FILE, 'w') as f:
            json.dump({"setup_done": True, "auth_token": token}, f)

        target_exe = sys.executable
        is_frozen = getattr(sys, 'frozen', False)

        if move_var.get() and OS_TYPE == "Windows" and is_frozen:
            dest_path = os.path.join(APP_DIR, os.path.basename(target_exe))
            if target_exe != dest_path:
                try:
                    shutil.copy2(target_exe, dest_path)
                    target_exe = dest_path
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to copy file: {e}")
                    return

        if startup_var.get():
            add_to_startup(target_exe)

        # Relaunch from new location if moved
        if os.path.abspath(target_exe) != os.path.abspath(sys.executable):
            try:
                subprocess.Popen([target_exe], close_fds=True)
                os._exit(0)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to relaunch: {e}")
                return

        root.destroy()

    tk.Button(root, text="Done", command=on_save, bg="#f9f9f9", fg="black", font=("Arial", 10, "bold"), width=15, borderwidth=0).pack(pady=20)
    root.mainloop()

# --- System Tray ---
def get_system_theme():
    try:
        if OS_TYPE == "Windows":
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
            value, _ = winreg.QueryValueEx(key, "SystemUsesLightTheme")
            return "light" if value == 1 else "dark"
        elif OS_TYPE == "Darwin":
            result = subprocess.run(['defaults', 'read', '-g', 'AppleInterfaceStyle'], capture_output=True, text=True)
            return "dark" if "Dark" in result.stdout else "light"
    except Exception:
        pass
    return "light" # Default fallback

def run_tray():
    if 'pystray' not in sys.modules:
        return

    def get_icon_image():
        theme = get_system_theme()
        icon_file = "monodark.png" if theme == "dark" else "monolight.png"

        paths = [
            resource_path(os.path.join('assets','icn', icon_file)),
            resource_path(os.path.join('assets','icn','default.png'))
        ]

        for path in paths:
            if os.path.exists(path):
                try:
                    return Image.open(path)
                except: pass

        return Image.new('RGBA', (64,64), (0,0,0,0))

    def on_show_ip(icon, item):
        if tk:
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
            token = config.get("auth_token")
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("SysLink Security Info", 
                                f"Host IP: {get_local_ip()}\n\n"
                                f"PIN: {token}\n\n"
                                "Enter these in Polygol Settings > Connections > Polygol Desktop to connect.")
            root.destroy()
            
    def on_exit(icon, item):
        icon.stop()
        os._exit(0) 

    def on_about(icon, item):
        if tk:
            root = tk.Tk()
            root.withdraw()

            def open_site():
                import webbrowser
                webbrowser.open("https://kirbindustries.github.io/polygol/syslink")

            about = tk.Toplevel()
            about.title("About")
            about.geometry("300x180")
            about.configure(bg="#1c1c1c")

            tk.Label(about, text="Polygol Desktop Companion", bg="#1c1c1c", fg="#ffffff", font=("Arial", 12, "bold")).pack(pady=10)
            tk.Label(about, text=f"Ver {VERSION}.", bg="#1c1c1c", fg="#b3b3b3").pack(pady=5)
            tk.Label(about, text="This product and Polygol is made by kirbIndustries", bg="#1c1c1c", fg="#ffffff").pack(pady=5)

            tk.Button(about, text="Open website", command=open_site).pack(pady=10)

            about.mainloop()

    image = get_icon_image()
    menu = pystray.Menu(
        pystray.MenuItem(f'IP: {get_local_ip()}', lambda: None),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('Connection Info...', on_show_ip),
        pystray.MenuItem('About...', on_about),
        pystray.MenuItem('Exit', on_exit)
    )
    icon = pystray.Icon("Polygol Desktop Companion", image, "Polygol Desktop Companion", menu)
    icon.run()

async def main():
    print(f"Starting PolygolDesktopCompanion on {OS_TYPE}...")
    print(f"Listening on ws://0.0.0.0:19420 (Local IP: {get_local_ip()})")
    async with websockets.serve(handler, "0.0.0.0", 19420):
        await asyncio.Future()

def start_background_loop():
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.run_until_complete(main())

def enforce_single_instance(port=19421):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", port))
        return s
    except OSError:
        print("Another instance is already running.")
        sys.exit(0)

if __name__ == "__main__":
    instance_lock = enforce_single_instance()
    if not os.path.exists(CONFIG_FILE):
        run_setup()

    # Start WebSocket Server in Background Thread
    server_thread = threading.Thread(target=start_background_loop, daemon=True)
    server_thread.start()

    # Start System Tray in Main Thread
    run_tray()

    # Fallback block if Tray fails to load
    while True:
        import time
        time.sleep(1000)