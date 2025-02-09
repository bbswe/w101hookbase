import asyncio
import time
import win32gui
import win32process
import psutil
import keyboard
import threading
from wizwalker.client import Client

exit_flag = False  # Global flag to signal script exit

def detect_ctrl_k():
    """Continuously checks for CTRL+K and exits the script when detected."""
    global exit_flag
    while not exit_flag:
        if keyboard.is_pressed("ctrl+k"):
            print("\n🔴 CTRL+K detected! Exiting safely...")
            exit_flag = True
            break
        time.sleep(0.1)

ctrl_k_thread = threading.Thread(target=detect_ctrl_k, daemon=True)
ctrl_k_thread.start()

def get_active_wizard101_window(timeout=10):
    """Waits for the user to select a valid Wizard101 window within a given time."""
    print(f"🖱️ Click on the Wizard101 window to select it. Waiting for up to {timeout} seconds...")
    start_time = time.time()
    while time.time() - start_time < timeout and not exit_flag:
        hwnd = win32gui.GetForegroundWindow()
        if hwnd:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                process = psutil.Process(pid)
                if process.name().lower() == "WizardGraphicalClient.exe":
                    print(f"✅ Selected active Wizard101 process: {process.name()} (PID: {pid})")
                    return hwnd
                else:
                    print(f"❌ Selected process is not Wizard101: {process.name()}")
            except psutil.NoSuchProcess:
                print("❌ Process no longer exists.")
        time.sleep(1)
    print("❌ No valid Wizard101 window selected within the time limit.")
    return None

async def cleanup_hooks(client):
    """Unhooks all hooks when closing the script."""
    if client:
        try:
            print("\n🔄 Cleaning up hooks before exiting...")
            await client.hook_handler.deactivate_player_hook()
            await client.hook_handler.close()
            print("✅ All hooks successfully removed!")
        except Exception as e:
            print(f"⚠️ Failed to unhook properly: {e}")

async def get_client():
    """Attaches to the Wizard101 client and verifies player existence."""
    window_handle = get_active_wizard101_window(timeout=10)
    if window_handle is None or exit_flag:
        print("❌ No valid Wizard101 window selected. Please click on the game window and try again.")
        return None

    client = Client(window_handle=window_handle)
    print(f"✅ Attached to Wizard101 client: {client}")

    if client.hook_handler.is_running():
        print("⚠️ Hooks are still running, deactivating first...")
        try:
            await client.hook_handler.deactivate_client_hook()
            await asyncio.sleep(2)
            print("✅ Hooks deactivated successfully!")
        except Exception as e:
            print(f"⚠️ No active hooks to deactivate: {e}")

    print("🔄 Resetting hook handler before rehooking...")
    await client.hook_handler.close()
    await asyncio.sleep(2)

    try:
        print("🔄 Activating player hook...")
        await asyncio.sleep(1)
        await client.hook_handler.activate_player_hook()
        print("✅ Player hook activated!")
    except Exception as e:
        print(f"❌ Failed to activate player hook: {e}")

    for attempt in range(20):
        if exit_flag:
            return None
        try:
            base_address = await client.body.read_base_address()
            print(f"🔍 Player Base Address: {hex(base_address)}")
            if base_address == 0:
                print("⏳ Attempting to get player base address again...")
            else:
                print("✅ Player base address found!")
                return client
        except Exception as e:
            print(f"⚠️ Attempt {attempt + 1}/20: Error reading player base address - {e}")
        await asyncio.sleep(1)
    print("❌ Player is not active. Exiting.")
    return None

async def print_positions_continuously(client):
    """
    Continuously prints the player's current position and yaw until the exit flag is set.
    """
    try:
        while not exit_flag:
            pos = await client.body.position()
            yaw = await client.body.yaw()
            print(f"📍 Position: X={pos.x:.2f}, Y={pos.y:.2f}, Yaw={yaw:.2f}")
            await asyncio.sleep(1)  # Update every second
    except Exception as e:
        print(f"⚠️ Error while printing position: {e}")

async def main():
    """Main function to attach to Wizard101 and continuously print the player's position."""
    client = await get_client()
    if client:
        print("✅ Hook is working! Continuously printing player's position...")
        await print_positions_continuously(client)
        await cleanup_hooks(client)
    else:
        print("No valid client attached.")

if __name__ == '__main__':
    asyncio.run(main())
