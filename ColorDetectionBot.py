import tkinter as tk
from tkinter import Toplevel, Canvas, colorchooser
import mss
import cv2
import numpy as np
import threading
import win32gui
import time
from pywinauto.keyboard import send_keys

# Constant for the Window Title
WINDOW_TITLE = "Window Title"
# Constant for the key to be pressed
KEY_TO_PRESS = '^'
# Constant for the color bounds, so if it is similar to the selected color by this much it will be detected
COLOR_BOUNDS = 10
# Constant for the delay after pressing
PRESS_DELAY = 5.0 # seconds

class AreaSelector:
    """
    A class to facilitate the selection of an area on the screen. This class creates a
    transparent window overlay that allows the user to select a rectangular area.
    """
    def __init__(self, root, window_rect):
        """
        Initializes the Area Selector window.

        Parameters:
        - root: The parent Tkinter object.
        - window_rect: A tuple containing the (x, y, width, height) of the target window.
        """
        # Window setup
        self.root = root
        self.top = Toplevel(root)
        self.top.attributes("-topmost", True)
        self.canvas = Canvas(self.top, cursor="cross")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Set geometry and transparency
        x, y, w, h = window_rect
        self.top.geometry(f"{w}x{h}+{x}+{y}")
        self.top.overrideredirect(True)
        self.top.wait_visibility(self.top)
        self.top.wm_attributes('-alpha', 0.3)

        # Bind mouse events
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

        # Initialize selection attributes
        self.start_x = None
        self.start_y = None
        self.rect = None
        self.selection = None

    # Event handlers for mouse actions
    def on_press(self, event):
        # Handle initial mouse button press
        self.start_x = event.x
        self.start_y = event.y
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red')

    def on_drag(self, event):
        # Handle mouse drag event to update the rectangle
        end_x, end_y = event.x, event.y
        self.canvas.coords(self.rect, self.start_x, self.start_y, end_x, end_y)

    def on_release(self, event):
        # Finalize the selection on mouse button release
        self.selection = self.canvas.coords(self.rect)
        self.top.destroy()

    def get_area(self):
        # Returns the selected area coordinates
        while not self.selection:
            self.top.update()
        return self.selection

class ColorDetectionBot:
    def __init__(self, root):
        """
        Initializes the Detection Bot application.

        Parameters:
        - root: The root Tkinter window object.
        """
        # Initialize main application window and default settings
        self.root = root
        # Set initial window size
        self.root.geometry("220x370")
        # Disable resizing and maximize button
        self.root.resizable(False, False)
        self.root.attributes("-toolwindow", True)  # This removes the maximize button
        self.coordinates = (0, 0, 100, 100)  # Default coordinates for detection area
        self.color_to_detect = [0, 0, 0]     # Default color to detect (black)
        self.monitoring = False              # Flag to control monitoring state
        self.key_press_count = 0            # Counter for key presses
        self.key_press_condition = tk.IntVar(value=0)  # 0 for color detected, 1 for color not detected
        self.bot_status = tk.StringVar(value="Bot Status: Not Running")
        self.detection_status = tk.StringVar(value="Color not detected")
        self.detection_total = tk.StringVar(value="Times key has been pressed: 0")
        self.setup_ui()                      # Setup the user interface

    def setup_ui(self):
        """
        Sets up the user interface of the application.
        """
        self.root.title("Color Detection Bot")

        # UI setup for coordinate selection
        self.step1_label = tk.Label(self.root, text="Step 1: Select Detection Area")
        self.step1_label.pack()
        self.pick_coords_button = tk.Button(self.root, text="Select Area", command=self.pick_coordinates)
        self.pick_coords_button.pack()

        # UI separator
        tk.Label(self.root, text="").pack()

        # UI setup for color selection
        self.step2_label = tk.Label(self.root, text="Step 2: Pick Color to Detect")
        self.step2_label.pack()
        self.color_picker_button = tk.Button(self.root, text="Pick Color", command=self.pick_color)
        self.color_picker_button.pack()

        # UI separator
        tk.Label(self.root, text="").pack()

        # UI setup for key press condition
        self.step3_label = tk.Label(self.root, text="Step 3: Key Press Condition")
        self.step3_label.pack()

        tk.Radiobutton(self.root, text="Press key when color is detected", variable=self.key_press_condition, value=0).pack()
        tk.Radiobutton(self.root, text="Press key when color is not detected", variable=self.key_press_condition, value=1).pack()

        # UI separator
        tk.Label(self.root, text="").pack()

        # UI setup for bot control
        self.step4_label = tk.Label(self.root, text="Step 4: Start/Stop Bot")
        self.step4_label.pack()
        self.start_button = tk.Button(self.root, text="Start Bot", command=self.start_bot)
        self.start_button.pack()
        self.stop_button = tk.Button(self.root, text="Stop Bot", state=tk.DISABLED, command=self.stop_bot)
        self.stop_button.pack()

        # UI for status labels
        bot_status_label = tk.Label(self.root, textvariable=self.bot_status)
        bot_status_label.pack()
        detection_status_label = tk.Label(self.root, textvariable=self.detection_status)
        detection_status_label.pack()
        detection_total_label = tk.Label(self.root, textvariable=self.detection_total)
        detection_total_label.pack()

    def pick_coordinates(self):
        """
        Handles the selection of screen coordinates for the detection area.
        """
        # Find the target window and get its coordinates
        hwnd = win32gui.FindWindow(None, "Endless Online")
        if not hwnd:
            print("Target window not found!")
            return

        print("Target window found!")
        rect = win32gui.GetWindowRect(hwnd)
        x, y, w, h = rect
        w, h = w - x, h - y

        # Create and use the area selector to get coordinates
        selector = AreaSelector(self.root, window_rect=(x, y, w, h))
        print("Area selector defined.")
        self.coordinates = selector.get_area()
        print("Selected Coordinates:", self.coordinates)
        if self.coordinates:
            self.detection_status.set(f"Selected Area: {self.coordinates}")

    def pick_color(self):
        """
        Handles the selection of the color to be detected.
        """
        # Open a color picker dialog and update the selected color
        color_code = colorchooser.askcolor(title="Choose a color")
        if color_code:
            self.color_to_detect = [int(c) for c in color_code[0]]
            self.detection_status.set(f"Selected Color: {self.color_to_detect}")

    def start_bot(self):
        """
        Starts the color detection bot.
        """
        # Start monitoring in a separate thread
        self.monitoring = True
        self.thread = threading.Thread(target=self.monitor_color, daemon=True)
        self.thread.start()
        # Update UI to reflect running state
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.bot_status.set("Bot Status: Running")

    def stop_bot(self):
        """
        Stops the color detection bot.
        """
        # Stop the monitoring thread
        self.monitoring = False
        try:
            self.thread.join()
        except RuntimeError as e:
            print(f"Error while stopping the thread: {e}")
        # Update UI to reflect stopped state
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.bot_status.set("Bot Status: Not Running")


    def monitor_color(self):
        """
        Monitors the selected area for the specified color.
        """
        # Find the target window using the WINDOW_TITLE constant
        hwnd = win32gui.FindWindow(None, WINDOW_TITLE)
        if not hwnd:
            print(f"Target window '{WINDOW_TITLE}' not found!")
            return

        window_rect = win32gui.GetWindowRect(hwnd)
        window_x, window_y = window_rect[0], window_rect[1]

        with mss.mss() as sct:
            monitor = {
                "top": int(window_y + self.coordinates[1]),
                "left": int(window_x + self.coordinates[0]),
                "width": int(self.coordinates[2] - self.coordinates[0]),
                "height": int(self.coordinates[3] - self.coordinates[1]),
            }

            while self.monitoring:
                screenshot = sct.grab(monitor)
                img = np.array(screenshot)

                img = cv2.cvtColor(img, cv2.COLOR_BGRA2RGB)

                color_detected = self.is_color_detected(img, self.color_to_detect)

                should_press_key = (
                    (color_detected and self.key_press_condition.get() == 0) or
                    (not color_detected and self.key_press_condition.get() == 1)
                )

                if should_press_key:
                    self.send_key_to_window(hwnd)
                    self.key_press_count += 1
                    self.detection_total.set(f"Times key has been pressed: {self.key_press_count}")
                    time.sleep(PRESS_DELAY)  # Time to wait when key has been pressed

                # Update status based on color detection
                if color_detected:
                    self.detection_status.set("Color detected")
                else:
                    self.detection_status.set("Color not detected")

                time.sleep(1.0)  # Time to wait between loop cycles
                    

    def is_color_detected(self, img, color_to_detect):
        """
        Checks if the specified color is detected in the given image.

        Parameters:
        - img: The image in which to detect the color.
        - color_to_detect: The RGB color to detect.
        """
        target_color = np.array(color_to_detect)
        lower_bound = target_color - COLOR_BOUNDS
        upper_bound = target_color + COLOR_BOUNDS
        mask = cv2.inRange(img, lower_bound, upper_bound)
        return np.any(mask)

    def send_key_to_window(self, hwnd):
        """
        Sends a key press to the specified window.

        Parameters:
        - hwnd: Handle to the window where the key will be sent.
        """
        # Save the handle of the currently focused window
        current_focused_window = win32gui.GetForegroundWindow()

        # Set the target window as foreground
        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.01)
            send_keys(KEY_TO_PRESS)  # Simulate key press
            time.sleep(0.01)
            send_keys(KEY_TO_PRESS)  # Release key

            # Restore focus to the previously focused window
            win32gui.SetForegroundWindow(current_focused_window)
        except:
            print(f"Error while focusing the window")

# Main application execution
if __name__ == "__main__":
    root = tk.Tk()
    app = ColorDetectionBot(root)
    root.mainloop()