import subprocess
import sys

def ppt_control_module():
    print("PowerPoint control module is now working.")
    try:
        # Replace 'python' with the full path to the Python executable of your environment
        python_executable = sys.executable
        subprocess.run([python_executable, "pptControl.py"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"An error occurred while running pptControl.py: {e}")

def virtual_whiteboard_module():
    print("Virtual whiteboard module is now working.")
    try:
        # Replace 'python' with the full path to the Python executable of your environment
        python_executable = sys.executable
        subprocess.run([python_executable, "aiVirtualWhiteboard.py"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"An error occurred while running pptControl.py: {e}")

def invalid_input():
    print("Invalid input. Please press 'A' for PPT control or 'B' for Virtual Whiteboard.")

# Dictionary to map keys to functions
switch = {
    'A': ppt_control_module,
    'B': virtual_whiteboard_module
}

# Function to handle key press
def main():
    while True:
        key = input("Press 'A' for PPT control, 'B' for AI Virtual Whiteboard, or 'Q' to exit: ").upper()
        # Call the appropriate function based on the key press
        switch.get(key, invalid_input)()
        if key == 'Q':
            print("Exiting the program.")
            break




if __name__ == "__main__":
    main()

