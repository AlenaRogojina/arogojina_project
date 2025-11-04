"""
Simple demo of global screen color picking
Click the button, then click anywhere on your screen to pick a color
"""

import tkinter as tk
from PIL import Image, ImageTk
import pyautogui
from pynput import mouse

class ColorPickerDemo:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Screen Color Picker Demo")
        self.root.geometry("400x400")
        
        self.picking = False
        self.mouse_listener = None
        
        # Create UI elements
        self.setup_ui()
        
    def setup_ui(self):
        # Instructions
        instructions = tk.Label(self.root, 
                               text="Click 'Pick Color' then click anywhere on screen\nto capture that pixel's color",
                               font=("Arial", 10),
                               justify="center")
        instructions.pack(pady=10)
        
        # Pick button
        self.pick_btn = tk.Button(self.root, 
                                 text="Pick Screen Color", 
                                 command=self.toggle_picking,
                                 font=("Arial", 12),
                                 bg="lightblue")
        self.pick_btn.pack(pady=10)
        
        # Status label
        self.status = tk.Label(self.root, text="Ready to pick", font=("Arial", 10))
        self.status.pack(pady=5)
        
        # Color display
        self.color_frame = tk.Frame(self.root, width=200, height=100, bg="white", relief="solid", bd=2)
        self.color_frame.pack(pady=20)
        self.color_frame.pack_propagate(False)
        
        # Color info label
        self.color_info = tk.Label(self.root, text="No color picked yet", font=("Arial", 10))
        self.color_info.pack(pady=5)
        
    def toggle_picking(self):
        if not self.picking:
            self.start_picking()
        else:
            self.stop_picking()
            
    def start_picking(self):
        self.picking = True
        self.pick_btn.config(text="Cancel Picking", bg="red")
        self.status.config(text="Click anywhere on screen to pick color...")
        
        # Start global mouse listener
        self.mouse_listener = mouse.Listener(on_click=self.on_screen_click)
        self.mouse_listener.start()
        
    def stop_picking(self):
        self.picking = False
        self.pick_btn.config(text="Pick Screen Color", bg="lightblue")
        self.status.config(text="Ready to pick")
        
        if self.mouse_listener:
            self.mouse_listener.stop()
            self.mouse_listener = None
            
    def on_screen_click(self, x, y, button, pressed):
        # Only respond to left mouse button press
        if pressed and button == mouse.Button.left and self.picking:
            try:
                # Capture screen and get pixel color
                screenshot = pyautogui.screenshot()
                r, g, b = screenshot.getpixel((x, y))
                
                # Update display
                self.update_color_display(r, g, b, x, y)
                
                # Stop picking
                self.stop_picking()
                
            except Exception as e:
                print(f"Error picking color: {e}")
                self.stop_picking()
                
    def update_color_display(self, r, g, b, x, y):
        # Convert RGB to hex
        hex_color = f"#{r:02x}{g:02x}{b:02x}"
        
        # Update color frame background
        self.color_frame.config(bg=hex_color)
        
        # Update info label
        info_text = f"RGB: ({r}, {g}, {b})\nHex: {hex_color}\nPosition: ({x}, {y})"
        self.color_info.config(text=info_text)
        
        # Update status
        self.status.config(text=f"Picked color: {hex_color}")
        
        print(f"Color picked at screen position ({x}, {y}): RGB({r}, {g}, {b}) = {hex_color}")
        
    def run(self):
        try:
            self.root.mainloop()
        finally:
            # Clean up mouse listener if still running
            if self.mouse_listener:
                self.mouse_listener.stop()

if __name__ == "__main__":
    print("Starting Screen Color Picker Demo...")
    print("Make sure you have installed: pip install pillow pyautogui pynput")
    
    demo = ColorPickerDemo()
    demo.run()