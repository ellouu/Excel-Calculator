import os
import time
import subprocess


def open_excel_calculator():
    excel_file = r" " # your excel file path inside the ""

    try:
        if os.path.exists(excel_file):
            subprocess.Popen(['start', 'excel', excel_file], shell=True)
            print(f"✓ Opened Excel calculator: {excel_file}")
        else:
            print(f"✗ Excel file not found at: {excel_file}")
            print("Please update the 'excel_file' path in the code.")
    except Exception as e:
        print(f"✗ Error opening Excel: {e}")


log_path = r" " # your log txt file path inside the "" (if you want one)


print("Excel Calculator Controller")
print("=" * 50)
open_excel_calculator()

print("\nWaiting for Excel to load (2 seconds)...")
time.sleep(2)

print("\nExcel Calculator Monitor")
print("Press buttons in Excel. Results will show when you press '='")
print("Press Ctrl+C to stop\n")

try:
    while True:
        try:
            with open(log_path, "r") as f:
                lines = f.readlines()
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                if line.startswith("=:"):
                    result = line[2:]  # Remove "=:" prefix
                    print(f"\n=== RESULT: {result} ===")
                else:
                    print(f"Button: {line}")
            with open(log_path, "w") as f:
                f.write("")

        except FileNotFoundError:
            pass
        except Exception as e:
            print(f"Error: {e}")

        time.sleep(0.1)

except KeyboardInterrupt:

    print("\nStopped monitoring.")
