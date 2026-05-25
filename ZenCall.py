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

ZCaller_logo = r"""

███████╗ ██████╗ █████╗ ██╗     ██╗     
╚══███╔╝██╔════╝██╔══██╗██║     ██║     
  ███╔╝ ██║     ███████║██║     ██║     
 ███╔╝  ██║     ██╔══██║██║     ██║     
███████╗╚██████╗██║  ██║███████╗███████╗
╚══════╝ ╚═════╝╚═╝  ╚═╝╚══════╝╚══════╝ LITE
-------------------------------------
     A u t o m a t e d  D i a l e r  for Zentrex
"""

def show_startup_logo():
    print(ZCaller_logo)  # Cyan color text
    print("Initializing ZCall Lite...\n")
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

# Function to switch to Zentrex
def switch_to_zentrex():
    try:
        zentrex_window = gw.getWindowsWithTitle("Zentrex")

        if zentrex_window:
            if zentrex_window[0].isMinimized:
                zentrex_window[0].restore()

            zentrex_window[0].activate()
            print("Switched to Zentrex.")

        else:
            print("Zentrex window not found.")

    except Exception as e:
        print(f"Error switching to Zentrex: {e}")

def switch_to_terminal():
    # Switch back to Python script terminal
    try:
        possible_titles = ["Python", "cmd", "Windows Terminal", "PowerShell","py.exe","ZCall Lite"]
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
    prev_number = number

    # Validate number
    if len(str(number)) != 8:
        print(f"Invalid number length ({number})")
        return

    switch_to_zentrex()

    time.sleep(0.5)

     # Navigate to textbox
    active_window = gw.getActiveWindow()

    if active_window:
        center_x = active_window.left + (active_window.width // 2)
        center_y = active_window.top + 80

        pyautogui.click(center_x, center_y)

    time.sleep(0.2)

    # Clear previous input
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')

    # Type phone number
    pyautogui.write(str(number), interval=0.03)

    time.sleep(0.3)

    # Press Enter to Call
    pyautogui.press('enter')

    print(f"Dialed: {number}")

    time.sleep(1)

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

def end_call():

    switch_to_zentrex()

    time.sleep(0.5)

    active_window = gw.getActiveWindow()

    if active_window:

        # Calculate dynamic button position
        end_x = active_window.left + active_window.width - 70

        # Near bottom of window
        end_y = active_window.top + active_window.height - 90

        pyautogui.click(end_x, end_y)

        print("Call ended.")

    else:
        print("Zentrex window not active.")

    time.sleep(0.5)

    switch_to_terminal()


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
    switch_to_zentrex()
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
            end_call()  # End call in BizPhone
            print("Call ended. Enter status:")
            continue  

        if call_status.lower() == "prev" and prev_number:
            print(f"Redialing: {prev_number}")
            switch_to_zentrex()
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
            print("Ending call in Zentrex...")
            end_call()
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
                switch_to_zentrex()
                pyautogui.press('0')  # End call in Zentrex
                time.sleep(1)

                print(f"Calling: {number}")
                switch_to_zentrex()
                dial_number(number)

                call_status = listen_for_status()
                
                if call_status:
                    while True:
                        print(call_status_str)
                        call_status = input().strip()
                            
                        if call_status.lower() == "prev" and prev_number:
                            print(f"Redialing: {prev_number}")
                            switch_to_zentrex()
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
print("Switch to Zentrex, press '.' to start dialing, or Esc to pause/resume.")
with keyboard.Listener(on_press=on_press) as listener:
    listener.join()
