import pyautogui
import openpyxl
import time
import os
import pygetwindow as gw
from pynput import keyboard
import queue
import json
import sys
import subprocess
import psutil
from datetime import datetime


# Define file paths
file_path = r"Calling.xlsx"
excel_path = r"C:\Program Files\Microsoft Office\Office16\EXCEL.EXE"

alice_logo = r"""
 █████  ██      ██  ██████  ███████  
██   ██ ██      ██ ██       ██       
███████ ██      ██ ██       █████    
██   ██ ██      ██ ██       ██       
██   ██ ███████ ██  ██████  ███████  LITE
-------------------------------------
     A u t o m a t e d  D i a l e r  
"""

def show_startup_logo():
    print(alice_logo)  # Cyan color text
    print("Initializing ALICE...\n")
    time.sleep(2)  # Pause for effect

show_startup_logo()

# Try to open the Excel file
try:
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    print("Excel file loaded successfully.")
except Exception as e:
    print(f"Error opening Excel file: {e}")
    subprocess.Popen([excel_path])  # Open Excel if not open
    time.sleep(3)  # Allow time to open
    exit()

# Control variables
current_row = 2  # Start from row 2 (assuming row 1 has headers)
paused = False
prev_number = None  # Track last dialed number
call_status_str = "Enter the call status (DNC, NPU, Not Interested, Follow Up, NIU, Appt, prev)"
call_next_str = "Press '.' to call the next number, '\\' to end call manually, or Esc to pause."

# Function to switch to BizPhone (LDPlayer)
def switch_to_bizphone():
    try:
        ldplayer_window = gw.getWindowsWithTitle("LDPlayer")
        if ldplayer_window:
            if ldplayer_window[0].isMinimized:
                ldplayer_window[0].restore()
            ldplayer_window[0].activate()
            print("Switched to BizPhone.")
        else:
            print("BizPhone window not found.")
    except Exception as e:
        print(f"Error switching to BizPhone: {e}")

def switch_to_terminal():
    # Switch back to Python script terminal
    try:
        possible_titles = ["Python", "cmd", "Windows Terminal", "PowerShell","py.exe"]
        python_windows = [w for w in gw.getAllWindows() if any(title in w.title for title in possible_titles)]

        if python_windows:
            python_windows[0].activate()
            print("Switched back to Python terminal.")

            # Clear any unwanted input (backspace multiple times)
            for _ in range(10):  # Adjust based on expected input length
                pyautogui.press('backspace')

            print("Ready for call status input.")
        else:
            print("Python terminal not found.")
    except Exception as e:
        print(f"Error switching back to Python terminal: {e}")

# Function to dial a number
def dial_number(number):
    global prev_number
    prev_number = number  # Save last dialed number

    pyautogui.press('0')  # End previous call
    time.sleep(1)

    # Check if the phone number has exactly 8 digits
    if len(str(number)) != 8:
        print(f"Invalid number length ({number}). Redialing...")
        return dial_number(number)  # Recursively retry dialing

    for _ in range(8):
        pyautogui.press('-')
    print("Cleared numpad.")

    for digit in str(number):
        pyautogui.press(digit)
    pyautogui.press('+')  # Press "+" instead of Enter
    print(f"Dialed: {number}")

    time.sleep(1)  # Allow some time for dialing

    # Switch back to Python script terminal
    switch_to_terminal()
        
def listen_for_status():
    print("Listening for status... (Simulating for now)")
    time.sleep(5)  # Simulate call listening
    return "Completed"  # Dummy return value for now

# Function to update Excel file
def update_excel(current_row, status, remarks=""):
    try:
        close_excel()
        
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active

        sheet[f"B{current_row}"] = status  # Update status column
        sheet[f"C{current_row}"] = datetime.now().strftime("%d/%m/%Y")  # Add today's date
        if status in ["Follow Up", "Appt"]:
            sheet[f"D{current_row}"] = remarks  # Add remarks column if applicable
        
        wb.save(file_path)
        wb.close()
        print(f"Row {current_row} updated with status '{status}', date, and remarks (if any).")

    except Exception as e:
        print(f"Error updating Excel: {e}")

# Function to close Excel processes
def close_excel():
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'] and 'excel' in process.info['name'].lower():
            try:
                process.terminate()
                print("Closed Excel process.")
                time.sleep(1)
            except Exception as e:
                print(f"Error closing Excel process: {e}")

# Function to handle calls
def handle_call():
    global current_row

    while sheet[f"B{current_row}"].value:  # Skip rows with existing status
        current_row += 1
    
    number = sheet[f"A{current_row}"].value
    if not number:
        print("No more numbers to call.")
        return

    print(f"Calling: {number}")
    switch_to_bizphone()
    dial_number(number)

    while True:
        call_status = input("Enter call status (DNC, NPU, Not Interested, Follow Up, NIU, Appt, prev, dc): ").strip()

        # Map shortcuts to full status names
        status_mapping = {
            "NI": "Not Interested",
            "FU": "Follow Up"
        }
        call_status = status_mapping.get(call_status, call_status)  # Convert shortcut to full status

        if call_status.lower() == "dc":
            pyautogui.press('0')  # End call in BizPhone
            print("Call ended. Enter status:")
            continue  

        if call_status.lower() == "prev" and prev_number:
            print(f"Redialing: {prev_number}")
            switch_to_bizphone()
            dial_number(prev_number)
            continue  

        if call_status in ["DNC", "NPU", "Not Interested", "Follow Up", "NIU", "Appt"]:
            remarks = ""
            if call_status in ["Follow Up", "Appt"]:
                remarks = input("Enter remarks: ").strip()

            update_excel(current_row, call_status, remarks)
            current_row += 1
            print("Call completed. Moving to next number.")
            break  

        print("Invalid status. Try again.")

# Function to handle keypress events
def on_press(key):
    global current_row, paused

    try:
        if key.char == '\\':  # Press "\" to manually end call
            print("Ending call in BizPhone...")
            switch_to_bizphone()
            pyautogui.press('0')  # End call in BizPhone
            switch_to_terminal()
            print("Call ended. You can now enter status or make another call.")

        elif key.char == '.':  # Press "." to dial numbers in sequence
            if paused:
                print("Script is paused. Press Esc to resume.")
                return

            while True:
                while sheet[f"B{current_row}"].value:  # Skip rows with existing status
                    current_row += 1
                
                number = sheet[f"A{current_row}"].value
                if not number:
                    print("No more numbers to call.")
                    break

                # End the previous call first
                print("Ending previous call before dialing the next number...")
                switch_to_bizphone()
                pyautogui.press('0')  # End call in BizPhone
                time.sleep(1)

                print(f"Calling: {number}")
                switch_to_bizphone()
                dial_number(number)

                call_status = listen_for_status()
                
                if call_status:
                    while True:
                        print(call_status_str)
                        call_status = input().strip()
                            
                        if call_status.lower() == "prev" and prev_number:
                            print(f"Redialing: {prev_number}")
                            switch_to_bizphone()
                            dial_number(prev_number)
                            continue  # Stay in the loop until a real status is entered

                        if call_status in ["DNC", "NPU", "Not Interested", "Follow Up", "NIU", "Appt"]:
                            remarks = ""
                            if call_status in ["Follow Up", "Appt"]:
                                print("Enter remarks:")
                                remarks = input().strip()
                            
                            update_excel(current_row, call_status, remarks)
                            current_row += 1
                            print("Call completed. Moving to next number.")
                            break  # Exit loop after getting a valid status

                print(call_next_str)
                return  # Wait for next key press

        elif key == keyboard.Key.esc:
            paused = not paused
            print("Script paused." if paused else "Script resumed.")

    except AttributeError:
        pass  # Handle special keys

# Start keyboard listener
print("Switch to BizPhone, press '.' to start dialing, or Esc to pause/resume.")
with keyboard.Listener(on_press=on_press) as listener:
    listener.join()
