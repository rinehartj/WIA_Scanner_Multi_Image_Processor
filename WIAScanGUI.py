import win32com.client
import pythoncom
from datetime import datetime
import os
import sys
import cv2
import numpy as np
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from PIL import Image, ImageTk
import json
import threading
import subprocess

# Get the correct path for bundled resources
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    BASE_DIR = sys._MEIPASS
else:
    # Running as script
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Settings file for persistence (in user directory, not in _MEIPASS)
if getattr(sys, 'frozen', False):
    SETTINGS_FILE = os.path.join(os.path.dirname(sys.executable), "scanner_settings.json")
else:
    SETTINGS_FILE = os.path.join(BASE_DIR, "scanner_settings.json")

EXIFTOOL_PATH = os.path.join(BASE_DIR, "tools", "exiftool.exe")

# WIA Constants
WIA_INTENT_NONE = 0x00000000
WIA_INTENT_IMAGE_TYPE_COLOR = 0x00000001
WIA_INTENT_IMAGE_TYPE_GRAYSCALE = 0x00000002
WIA_INTENT_IMAGE_TYPE_TEXT = 0x00000004


class CornerAdjustmentWindow:
    """Fullscreen window for adjusting corner points"""

    def __init__(self, parent, image, corners, callback):
        self.parent = parent
        self.original_image = image.copy()
        self.corners = corners.copy()
        self.callback = callback
        self.selected_corner = None
        self.dragging = False
        self.zoom_size = 150  # Size of zoom window
        self.zoom_factor = 3  # Zoom magnification

        # Create fullscreen window
        self.window = tk.Toplevel(parent)
        self.window.title("Adjust Corner Points")
        self.window.attributes('-fullscreen', True)
        self.window.configure(bg='black')

        # Get screen dimensions
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()

        # Reserve space for controls at bottom
        self.control_height = 120
        self.canvas_height = screen_height - self.control_height

        # Create canvas for main image
        self.canvas = tk.Canvas(self.window, bg='black', highlightthickness=0,
                                height=self.canvas_height)
        self.canvas.pack(fill=tk.BOTH, expand=False)

        # Create zoom canvas (initially hidden)
        self.zoom_canvas = tk.Canvas(self.window, bg='black', highlightthickness=2,
                                     highlightbackground='yellow',
                                     width=self.zoom_size, height=self.zoom_size)

        # Bind events
        self.canvas.bind('<Button-1>', self.on_mouse_down)
        self.canvas.bind('<B1-Motion>', self.on_mouse_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_mouse_up)
        self.canvas.bind('<Motion>', self.on_mouse_move)
        self.window.bind('<Escape>', lambda e: self.cancel())
        self.window.bind('<Return>', lambda e: self.apply())

        # Instructions frame at bottom
        self.create_instructions()

        # Wait for window to be drawn
        self.window.update()

        # Display image
        self.display_image()

    def create_instructions(self):
        """Create instruction panel at bottom"""
        frame = tk.Frame(self.window, bg='black', height=self.control_height)
        frame.pack(side=tk.BOTTOM, fill=tk.X)
        frame.pack_propagate(False)

        instructions = "Drag corner points to adjust • ENTER to apply • ESC to cancel"
        label = tk.Label(frame, text=instructions, bg='black', fg='white',
                         font=('Arial', 14, 'bold'))
        label.pack(pady=10)

        btn_frame = tk.Frame(frame, bg='black')
        btn_frame.pack()

        apply_btn = tk.Button(btn_frame, text="Apply", command=self.apply,
                              bg='green', fg='white', font=('Arial', 12, 'bold'),
                              padx=20, pady=10)
        apply_btn.pack(side=tk.LEFT, padx=10)

        cancel_btn = tk.Button(btn_frame, text="Cancel", command=self.cancel,
                               bg='red', fg='white', font=('Arial', 12, 'bold'),
                               padx=20, pady=10)
        cancel_btn.pack(side=tk.LEFT, padx=10)

    def display_image(self):
        """Display image with corner points"""
        # Get canvas size
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas_height

        # Create image with corners drawn
        display_img = self.original_image.copy()

        # Scale corners to display coordinates
        img_h, img_w = self.original_image.shape[:2]
        scale = min(canvas_width / img_w, canvas_height / img_h)

        # Ensure we don't scale up
        if scale > 1.0:
            scale = 1.0

        new_w = int(img_w * scale)
        new_h = int(img_h * scale)

        # Store scale and offset for coordinate conversion
        self.scale = scale
        self.offset_x = (canvas_width - new_w) // 2
        self.offset_y = (canvas_height - new_h) // 2

        # Draw lines between corners
        for i in range(4):
            pt1 = tuple(self.corners[i].astype(int))
            pt2 = tuple(self.corners[(i + 1) % 4].astype(int))
            cv2.line(display_img, pt1, pt2, (0, 255, 0), 4)

        # Draw corner points (larger and more visible - doubled size)
        for i, pt in enumerate(self.corners):
            pt_int = tuple(pt.astype(int))
            # Draw outer circle (doubled from 25 to 50)
            cv2.circle(display_img, pt_int, 50, (0, 255, 0), -1)
            # Draw inner circle (doubled from 20 to 40)
            cv2.circle(display_img, pt_int, 40, (255, 255, 255), -1)
            # Draw number
            cv2.putText(display_img, str(i + 1), pt_int,
                        cv2.FONT_HERSHEY_SIMPLEX, 2.5, (0, 0, 0), 6)

        # Resize image
        display_img = cv2.resize(display_img, (new_w, new_h))

        # Convert to RGB
        display_img = cv2.cvtColor(display_img, cv2.COLOR_BGR2RGB)

        # Convert to PhotoImage
        img_pil = Image.fromarray(display_img)
        self.photo = ImageTk.PhotoImage(img_pil)

        # Display on canvas
        self.canvas.delete("all")
        self.canvas.create_image(self.offset_x, self.offset_y,
                                 image=self.photo, anchor='nw')

    def show_zoom(self, canvas_x, canvas_y):
        """Show zoomed-in view around cursor"""
        img_x, img_y = self.canvas_to_image_coords(canvas_x, canvas_y)

        # Get region around cursor
        img_h, img_w = self.original_image.shape[:2]
        half_size = self.zoom_size // (2 * self.zoom_factor)

        x1 = max(0, int(img_x - half_size))
        y1 = max(0, int(img_y - half_size))
        x2 = min(img_w, int(img_x + half_size))
        y2 = min(img_h, int(img_y + half_size))

        if x2 - x1 > 0 and y2 - y1 > 0:
            # Extract region
            region = self.original_image[y1:y2, x1:x2].copy()

            # Draw crosshair at center
            center_x = int(img_x - x1)
            center_y = int(img_y - y1)
            cv2.drawMarker(region, (center_x, center_y), (255, 0, 0),
                           cv2.MARKER_CROSS, 20, 2)

            # Zoom the region
            zoom_h = (y2 - y1) * self.zoom_factor
            zoom_w = (x2 - x1) * self.zoom_factor
            zoomed = cv2.resize(region, (zoom_w, zoom_h), interpolation=cv2.INTER_NEAREST)

            # Crop to fit zoom window
            crop_h = min(self.zoom_size, zoom_h)
            crop_w = min(self.zoom_size, zoom_w)
            start_h = (zoom_h - crop_h) // 2
            start_w = (zoom_w - crop_w) // 2
            zoomed = zoomed[start_h:start_h + crop_h, start_w:start_w + crop_w]

            # Convert to RGB
            zoomed = cv2.cvtColor(zoomed, cv2.COLOR_BGR2RGB)
            img_pil = Image.fromarray(zoomed)
            self.zoom_photo = ImageTk.PhotoImage(img_pil)

            # Position zoom window near cursor but not blocking it
            zoom_x = canvas_x + 30
            zoom_y = canvas_y + 30

            # Keep zoom window on screen
            canvas_width = self.canvas.winfo_width()
            if zoom_x + self.zoom_size > canvas_width:
                zoom_x = canvas_x - self.zoom_size - 30
            if zoom_y + self.zoom_size > self.canvas_height:
                zoom_y = canvas_y - self.zoom_size - 30

            # Place zoom canvas
            self.zoom_canvas.place(x=zoom_x, y=zoom_y)
            self.zoom_canvas.delete("all")
            self.zoom_canvas.create_image(0, 0, image=self.zoom_photo, anchor='nw')

    def hide_zoom(self):
        """Hide zoom window"""
        self.zoom_canvas.place_forget()

    def canvas_to_image_coords(self, canvas_x, canvas_y):
        """Convert canvas coordinates to image coordinates"""
        img_x = (canvas_x - self.offset_x) / self.scale
        img_y = (canvas_y - self.offset_y) / self.scale
        return img_x, img_y

    def on_mouse_move(self, event):
        """Handle mouse movement"""
        if self.dragging:
            self.show_zoom(event.x, event.y)

    def on_mouse_down(self, event):
        """Handle mouse down event"""
        img_x, img_y = self.canvas_to_image_coords(event.x, event.y)

        # Find nearest corner
        min_dist = float('inf')
        closest_corner = None

        for i, pt in enumerate(self.corners):
            dist = np.sqrt((pt[0] - img_x) ** 2 + (pt[1] - img_y) ** 2)
            if dist < min_dist:
                min_dist = dist
                closest_corner = i

        # Select corner if close enough (100 pixels in image space - doubled from 50)
        if min_dist < 100:
            self.selected_corner = closest_corner
            self.dragging = True
            self.show_zoom(event.x, event.y)

    def on_mouse_drag(self, event):
        """Handle mouse drag event"""
        if self.dragging and self.selected_corner is not None:
            img_x, img_y = self.canvas_to_image_coords(event.x, event.y)

            # Clamp to image bounds
            img_h, img_w = self.original_image.shape[:2]
            img_x = max(0, min(img_w - 1, img_x))
            img_y = max(0, min(img_h - 1, img_y))

            # Update corner position
            self.corners[self.selected_corner] = [img_x, img_y]

            # Redraw
            self.display_image()
            self.show_zoom(event.x, event.y)

    def on_mouse_up(self, event):
        """Handle mouse up event"""
        self.dragging = False
        self.selected_corner = None
        self.hide_zoom()

    def apply(self):
        """Apply changes and close window"""
        self.callback(self.corners)
        self.window.destroy()

    def cancel(self):
        """Cancel changes and close window"""
        self.window.destroy()


class ScannerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Scanner Controller")
        self.root.geometry("1200x750")

        # Scanner variables
        self.scanner = None
        self.tiff_filepath = None
        self.cropped_images = []
        self.current_preview_index = 0
        self.full_scan_image = None  # Store full scan for manual add

        # Load settings
        self.load_settings()

        # Create GUI
        self.create_widgets()

    def load_settings(self):
        """Load persistent settings from file"""
        default_settings = {
            "dpi": 800,
            "scan_distance": 3.0,
            "jpg_quality": 98,
            "crop_pixels": 5,
            "save_tiff": True,
            "document_name": "Document",
            "output_folder": "Scans",
            "exif_date": "",
            "exif_title": ""
        }

        try:
            if os.path.exists(SETTINGS_FILE):
                with open(SETTINGS_FILE, 'r') as f:
                    self.settings = json.load(f)
            else:
                self.settings = default_settings
        except:
            self.settings = default_settings

    def save_settings(self):
        """Save settings to file"""
        try:
            with open(SETTINGS_FILE, 'w') as f:
                json.dump(self.settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def create_widgets(self):
        """Create all GUI widgets"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Configure grid weights for responsive layout
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=2)
        main_frame.rowconfigure(2, weight=1)

        # Left column container
        left_frame = ttk.Frame(main_frame)
        left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))

        # Settings Frame
        settings_frame = ttk.LabelFrame(left_frame, text="Scan Settings", padding="10")
        settings_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)

        # DPI
        ttk.Label(settings_frame, text="DPI:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.dpi_var = tk.IntVar(value=self.settings.get("dpi", 800))
        dpi_spin = ttk.Spinbox(settings_frame, from_=150, to=2400, increment=50,
                               textvariable=self.dpi_var, width=10)
        dpi_spin.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        # Scan Distance
        ttk.Label(settings_frame, text="Scan Distance (inches):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.distance_var = tk.DoubleVar(value=self.settings.get("scan_distance", 3.0))
        distance_spin = ttk.Spinbox(settings_frame, from_=1.0, to=12.0, increment=0.5,
                                    textvariable=self.distance_var, width=10)
        distance_spin.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        # JPG Quality
        ttk.Label(settings_frame, text="JPG Quality (%):").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.jpg_quality_var = tk.IntVar(value=self.settings.get("jpg_quality", 98))
        jpg_spin = ttk.Spinbox(settings_frame, from_=50, to=100, increment=1,
                               textvariable=self.jpg_quality_var, width=10)
        jpg_spin.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)

        # Crop Pixels
        ttk.Label(settings_frame, text="Crop Pixels:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.crop_pixels_var = tk.IntVar(value=self.settings.get("crop_pixels", 5))
        crop_spin = ttk.Spinbox(settings_frame, from_=0, to=50, increment=1,
                                textvariable=self.crop_pixels_var, width=10)
        crop_spin.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)

        # Save TIFF checkbox
        self.save_tiff_var = tk.BooleanVar(value=self.settings.get("save_tiff", True))
        tiff_check = ttk.Checkbutton(settings_frame, text="Save TIFF file",
                                     variable=self.save_tiff_var)
        tiff_check.grid(row=4, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        # Document Name
        ttk.Label(settings_frame, text="Document Name:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
        self.doc_name_var = tk.StringVar(value=self.settings.get("document_name", "Document"))
        doc_entry = ttk.Entry(settings_frame, textvariable=self.doc_name_var, width=20)
        doc_entry.grid(row=5, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        # Output Folder
        ttk.Label(settings_frame, text="Output Folder:").grid(row=6, column=0, sticky=tk.W, padx=5, pady=5)
        folder_subframe = ttk.Frame(settings_frame)
        folder_subframe.grid(row=6, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        self.folder_var = tk.StringVar(value=self.settings.get("output_folder", "Scans"))
        folder_entry = ttk.Entry(folder_subframe, textvariable=self.folder_var, width=20)
        folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        browse_btn = ttk.Button(folder_subframe, text="...", command=self.browse_folder, width=3)
        browse_btn.pack(side=tk.LEFT, padx=(5, 0))

        # EXIF Settings Frame
        exif_frame = ttk.LabelFrame(left_frame, text="EXIF Metadata (Optional)", padding="10")
        exif_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)

        # EXIF Date
        ttk.Label(exif_frame, text="Date (YYYY:MM:DD):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.exif_date_var = tk.StringVar(value=self.settings.get("exif_date", ""))
        exif_date_entry = ttk.Entry(exif_frame, textvariable=self.exif_date_var, width=20)
        exif_date_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        ttk.Label(exif_frame, text="(Leave blank for no date)",
                  font=('Arial', 8, 'italic')).grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=5)

        # EXIF Title
        ttk.Label(exif_frame, text="Title:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.exif_title_var = tk.StringVar(value=self.settings.get("exif_title", ""))
        exif_title_entry = ttk.Entry(exif_frame, textvariable=self.exif_title_var, width=20)
        exif_title_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        # Scan Button
        scan_btn = ttk.Button(left_frame, text="Scan", command=self.start_scan, width=20)
        scan_btn.grid(row=2, column=0, pady=10)

        # Progress/Log Frame
        log_frame = ttk.LabelFrame(left_frame, text="Status", padding="10")
        log_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        left_frame.rowconfigure(3, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, width=50, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Preview Frame (right side)
        preview_frame = ttk.LabelFrame(main_frame, text="Preview", padding="10")
        preview_frame.grid(row=0, column=1, rowspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        # Canvas for image display
        self.canvas = tk.Canvas(preview_frame, bg='gray', width=600, height=500)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Navigation and action buttons at bottom
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)

        self.prev_btn = ttk.Button(button_frame, text="← Previous", command=self.prev_image, state='disabled')
        self.prev_btn.pack(side=tk.LEFT, padx=5)

        self.image_label = ttk.Label(button_frame, text="No images")
        self.image_label.pack(side=tk.LEFT, padx=20)

        self.next_btn = ttk.Button(button_frame, text="Next →", command=self.next_image, state='disabled')
        self.next_btn.pack(side=tk.LEFT, padx=5)

        # Rotation buttons
        self.rotate_ccw_btn = ttk.Button(button_frame, text="↶ Rotate CCW",
                                         command=self.rotate_ccw, state='disabled')
        self.rotate_ccw_btn.pack(side=tk.LEFT, padx=5)

        self.rotate_cw_btn = ttk.Button(button_frame, text="↷ Rotate CW",
                                        command=self.rotate_cw, state='disabled')
        self.rotate_cw_btn.pack(side=tk.LEFT, padx=5)

        self.adjust_btn = ttk.Button(button_frame, text="Adjust Corners",
                                     command=self.open_adjustment_window, state='disabled')
        self.adjust_btn.pack(side=tk.LEFT, padx=20)

        self.add_btn = ttk.Button(button_frame, text="Add Image",
                                  command=self.add_manual_image, state='disabled')
        self.add_btn.pack(side=tk.LEFT, padx=5)

        self.remove_btn = ttk.Button(button_frame, text="Remove Image",
                                     command=self.remove_current_image, state='disabled')
        self.remove_btn.pack(side=tk.LEFT, padx=5)

        self.keep_btn = ttk.Button(button_frame, text="Keep All",
                                   command=self.keep_images, state='disabled')
        self.keep_btn.pack(side=tk.LEFT, padx=5)

        self.rescan_btn = ttk.Button(button_frame, text="Rescan",
                                     command=self.rescan, state='disabled')
        self.rescan_btn.pack(side=tk.LEFT, padx=5)

    def browse_folder(self):
        """Browse for output folder"""
        folder = filedialog.askdirectory(initialdir=self.folder_var.get(),
                                         title="Select Output Folder")
        if folder:
            self.folder_var.set(folder)
            self.settings["output_folder"] = folder
            self.save_settings()

    def log(self, message):
        """Add message to log"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()

    def start_scan(self):
        """Start scanning in a separate thread"""
        # Save current settings
        self.settings["dpi"] = self.dpi_var.get()
        self.settings["scan_distance"] = self.distance_var.get()
        self.settings["jpg_quality"] = self.jpg_quality_var.get()
        self.settings["crop_pixels"] = self.crop_pixels_var.get()
        self.settings["save_tiff"] = self.save_tiff_var.get()
        self.settings["document_name"] = self.doc_name_var.get()
        self.settings["output_folder"] = self.folder_var.get()
        self.settings["exif_date"] = self.exif_date_var.get()
        self.settings["exif_title"] = self.exif_title_var.get()
        self.save_settings()

        # Clear log and preview
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        self.canvas.delete("all")

        # Start scan in thread
        thread = threading.Thread(target=self.scan_thread)
        thread.daemon = True
        thread.start()

    def scan_thread(self):
        """Scan operation in separate thread"""
        try:
            # Initialize COM for this thread - MUST be first
            pythoncom.CoInitialize()

            self.log("Initializing scanner...")
            self.scanner = self.setup_scanner()

            if self.scanner is None:
                self.log("ERROR: Failed to initialize scanner")
                pythoncom.CoUninitialize()
                return

            self.log("Starting scan...")
            image = self.scan_document(self.scanner)

            if image is None:
                self.log("ERROR: Scan failed")
                pythoncom.CoUninitialize()
                return

            # Save TIFF if enabled
            if self.save_tiff_var.get():
                self.log("Saving TIFF file...")
                self.tiff_filepath = self.save_image(image)
                if self.tiff_filepath:
                    self.log(f"TIFF saved: {self.tiff_filepath}")
                else:
                    self.log("ERROR: Failed to save TIFF")
                    pythoncom.CoUninitialize()
                    return
            else:
                # Save temporary TIFF for processing in script directory
                self.log("Creating temporary file for processing...")
                if getattr(sys, 'frozen', False):
                    # When frozen, use executable directory
                    temp_dir = os.path.dirname(sys.executable)
                else:
                    # When running as script
                    temp_dir = os.path.dirname(os.path.abspath(__file__))
                self.tiff_filepath = os.path.join(temp_dir, "temp_scan.tif")
                image.SaveFile(self.tiff_filepath)

            # Store full scan image for manual add
            self.full_scan_image = cv2.imread(self.tiff_filepath)

            # Detect and crop documents
            self.log("Detecting documents...")
            self.cropped_images = self.detect_and_crop_documents(self.tiff_filepath)

            if self.cropped_images:
                self.log(f"Detected {len(self.cropped_images)} document(s)")
                self.current_preview_index = 0
                self.root.after(0, self.display_preview)
                self.root.after(0, self.enable_preview_controls)
            else:
                self.log("No documents detected - use 'Add Image' to manually add")
                self.root.after(0, self.enable_preview_controls)

            # Delete temp TIFF if not saving
            if not self.save_tiff_var.get() and os.path.exists(self.tiff_filepath):
                os.remove(self.tiff_filepath)

            # Uninitialize COM
            pythoncom.CoUninitialize()

        except Exception as e:
            self.log(f"ERROR: {str(e)}")
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def setup_scanner(self):
        """Initialize WIA and find the scanner."""
        try:
            device_manager = win32com.client.Dispatch("WIA.DeviceManager")
            scanner = None

            for i in range(1, device_manager.DeviceInfos.Count + 1):
                device_info = device_manager.DeviceInfos.Item(i)
                if "3200" in device_info.Properties("Name").Value or \
                        "Epson" in device_info.Properties("Name").Value:
                    scanner = device_info.Connect()
                    self.log(f"Found scanner: {device_info.Properties('Name').Value}")
                    break

            if scanner is None:
                if device_manager.DeviceInfos.Count > 0:
                    scanner = device_manager.DeviceInfos.Item(1).Connect()
                    self.log(f"Using scanner: {device_manager.DeviceInfos.Item(1).Properties('Name').Value}")

            return scanner
        except Exception as e:
            self.log(f"Error setting up scanner: {e}")
            return None

    def scan_document(self, scanner):
        """Perform the scan operation."""
        try:
            item = scanner.Items(1)

            # Set properties
            dpi = self.dpi_var.get()
            self.set_property_by_id(item, 6147, dpi, "Horizontal Resolution")
            self.set_property_by_id(item, 6148, dpi, "Vertical Resolution")
            self.set_property_by_id(item, 6146, WIA_INTENT_IMAGE_TYPE_COLOR, "Color Mode")

            # Set scan area based on distance
            distance = self.distance_var.get()
            extent = int(distance * dpi)
            self.set_property_by_id(item, 6151, extent, "X Extent")
            self.set_property_by_id(item, 6152, extent, "Y Extent")

            # Perform scan
            image = item.Transfer("{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}")
            self.log("Scan completed successfully")
            return image
        except Exception as e:
            self.log(f"Error during scan: {e}")
            return None

    def set_property_by_id(self, item, prop_id, value, prop_name="Unknown"):
        """Safely set a property by ID."""
        try:
            for i in range(1, item.Properties.Count + 1):
                prop = item.Properties(i)
                if prop.PropertyID == prop_id:
                    prop.Value = value
                    return True
            return False
        except:
            return False

    def save_image(self, image):
        """Save the scanned image as TIFF."""
        try:
            output_folder = self.settings.get("output_folder", "Scans")
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            date_str = datetime.now().strftime("%Y.%m.%d")
            doc_name = self.doc_name_var.get()
            filename = f"{date_str} {doc_name}.tif"
            filepath = os.path.join(output_folder, filename)

            counter = 1
            while os.path.exists(filepath):
                filename = f"{date_str} {doc_name}_{counter}.tif"
                filepath = os.path.join(output_folder, filename)
                counter += 1

            image.SaveFile(filepath)
            return filepath
        except Exception as e:
            self.log(f"Error saving image: {e}")
            return None

    def detect_and_crop_documents(self, image_path):
        """Detect and crop documents from scanned image"""
        image = cv2.imread(image_path)
        if image is None:
            return []

        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (7, 7), 0)
        thresh = cv2.adaptiveThreshold(blur, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                                       cv2.THRESH_BINARY_INV, blockSize=51, C=10)

        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (25, 25))
        closed = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)

        contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        h, w = gray.shape
        img_area = h * w
        min_area = img_area * 0.01

        cropped_data = []
        for cnt in contours:
            area = cv2.contourArea(cnt)
            if area < min_area:
                continue

            rect = cv2.minAreaRect(cnt)
            box = cv2.boxPoints(rect)
            box = np.float32(box)

            # Order the points properly
            box = self.order_points(box)

            cropped_data.append({
                'image': image.copy(),
                'corners': box,
                'original_corners': box.copy()
            })

        return cropped_data

    def rotate_cw(self):
        """Rotate current image 90 degrees clockwise"""
        if not self.cropped_images:
            return

        data = self.cropped_images[self.current_preview_index]

        # Rotate image 90 degrees clockwise
        data['image'] = cv2.rotate(data['image'], cv2.ROTATE_90_CLOCKWISE)

        # Rotate corners
        h, w = data['image'].shape[:2]
        new_corners = np.zeros_like(data['corners'])
        for i, (x, y) in enumerate(data['corners']):
            # Transform: (x, y) -> (h - y, x) for 90° CW rotation
            new_corners[i] = [h - y, x]

        data['corners'] = new_corners
        data['original_corners'] = new_corners.copy()

        self.display_preview()
        self.log(f"Rotated image {self.current_preview_index + 1} clockwise")

    def rotate_ccw(self):
        """Rotate current image 90 degrees counter-clockwise"""
        if not self.cropped_images:
            return

        data = self.cropped_images[self.current_preview_index]

        # Rotate image 90 degrees counter-clockwise
        data['image'] = cv2.rotate(data['image'], cv2.ROTATE_90_COUNTERCLOCKWISE)

        # Rotate corners
        h, w = data['image'].shape[:2]
        new_corners = np.zeros_like(data['corners'])
        for i, (x, y) in enumerate(data['corners']):
            # Transform: (x, y) -> (y, w - x) for 90° CCW rotation
            new_corners[i] = [y, w - x]

        data['corners'] = new_corners
        data['original_corners'] = new_corners.copy()

        self.display_preview()
        self.log(f"Rotated image {self.current_preview_index + 1} counter-clockwise")

    def add_manual_image(self):
        """Add manual image with full flatbed as default"""
        if self.full_scan_image is None:
            messagebox.showwarning("No Scan", "Please perform a scan first")
            return

        # Create default corners for full image
        h, w = self.full_scan_image.shape[:2]
        corners = np.array([
            [0, 0],
            [w - 1, 0],
            [w - 1, h - 1],
            [0, h - 1]
        ], dtype=np.float32)

        # Add to cropped images
        self.cropped_images.append({
            'image': self.full_scan_image.copy(),
            'corners': corners,
            'original_corners': corners.copy()
        })

        # Navigate to the new image
        self.current_preview_index = len(self.cropped_images) - 1
        self.display_preview()
        self.update_navigation_buttons()
        self.log(f"Added manual image (full flatbed) - Total: {len(self.cropped_images)} images")

    def remove_current_image(self):
        """Remove the current image from the list"""
        if not self.cropped_images:
            return

        if messagebox.askyesno("Remove Image",
                               f"Remove image {self.current_preview_index + 1} from the list?"):
            self.cropped_images.pop(self.current_preview_index)
            self.log(f"Removed image - Remaining: {len(self.cropped_images)} images")

            if self.cropped_images:
                # Adjust index if needed
                if self.current_preview_index >= len(self.cropped_images):
                    self.current_preview_index = len(self.cropped_images) - 1
                self.display_preview()
                self.update_navigation_buttons()
            else:
                # No images left
                self.canvas.delete("all")
                self.image_label.config(text="No images")
                self.reset_preview()

    def display_preview(self):
        """Display current cropped image"""
        if not self.cropped_images:
            return

        data = self.cropped_images[self.current_preview_index]
        image = data['image']
        corners = data['corners']

        # Apply transform
        warped = self.four_point_transform(image, corners)

        # Apply crop pixels
        crop_px = self.crop_pixels_var.get()
        if crop_px > 0:
            h, w = warped.shape[:2]
            if h > 2 * crop_px and w > 2 * crop_px:
                warped = warped[crop_px:h - crop_px, crop_px:w - crop_px]

        # Convert to RGB for display
        display_img = cv2.cvtColor(warped, cv2.COLOR_BGR2RGB)

        # Resize to fit canvas
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        h, w = display_img.shape[:2]

        scale = min(canvas_width / w, canvas_height / h, 1.0)
        new_w = int(w * scale)
        new_h = int(h * scale)

        display_img = cv2.resize(display_img, (new_w, new_h))

        # Convert to PhotoImage
        img_pil = Image.fromarray(display_img)
        self.photo = ImageTk.PhotoImage(img_pil)

        # Display on canvas
        self.canvas.delete("all")
        self.canvas.create_image(canvas_width // 2, canvas_height // 2, image=self.photo)

        # Update label
        self.image_label.config(text=f"Image {self.current_preview_index + 1} of {len(self.cropped_images)}")

    def open_adjustment_window(self):
        """Open fullscreen window for corner adjustment"""
        if not self.cropped_images:
            return

        data = self.cropped_images[self.current_preview_index]
        image = data['image']
        corners = data['corners']

        # Open adjustment window
        CornerAdjustmentWindow(self.root, image, corners, self.on_corners_adjusted)

    def on_corners_adjusted(self, new_corners):
        """Callback when corners are adjusted"""
        if self.cropped_images:
            self.cropped_images[self.current_preview_index]['corners'] = new_corners
            self.display_preview()
            self.log(f"Corners adjusted for image {self.current_preview_index + 1}")

    def four_point_transform(self, image, pts):
        """Apply perspective transform"""
        rect = self.order_points(pts)
        (tl, tr, br, bl) = rect

        widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
        widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
        maxWidth = max(int(widthA), int(widthB))

        heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
        heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
        maxHeight = max(int(heightA), int(heightB))

        dst = np.array([[0, 0], [maxWidth - 1, 0],
                        [maxWidth - 1, maxHeight - 1], [0, maxHeight - 1]], dtype="float32")

        M = cv2.getPerspectiveTransform(rect, dst)
        warped = cv2.warpPerspective(image, M, (maxWidth, maxHeight))

        return warped

    def order_points(self, pts):
        """Order points: top-left, top-right, bottom-right, bottom-left"""
        rect = np.zeros((4, 2), dtype="float32")
        s = pts.sum(axis=1)
        rect[0] = pts[np.argmin(s)]
        rect[2] = pts[np.argmax(s)]
        diff = np.diff(pts, axis=1)
        rect[1] = pts[np.argmin(diff)]
        rect[3] = pts[np.argmax(diff)]
        return rect

    def enable_preview_controls(self):
        """Enable preview control buttons"""
        self.update_navigation_buttons()
        self.rotate_cw_btn.config(state='normal')
        self.rotate_ccw_btn.config(state='normal')
        self.adjust_btn.config(state='normal')
        self.add_btn.config(state='normal')
        self.remove_btn.config(state='normal')
        self.keep_btn.config(state='normal')
        self.rescan_btn.config(state='normal')

    def update_navigation_buttons(self):
        """Update navigation button states"""
        if len(self.cropped_images) > 1:
            self.prev_btn.config(state='normal')
            self.next_btn.config(state='normal')
        else:
            self.prev_btn.config(state='disabled')
            self.next_btn.config(state='disabled')

    def prev_image(self):
        """Show previous image"""
        if self.current_preview_index > 0:
            self.current_preview_index -= 1
            self.display_preview()

    def next_image(self):
        """Show next image"""
        if self.current_preview_index < len(self.cropped_images) - 1:
            self.current_preview_index += 1
            self.display_preview()

    def keep_images(self):
        """Save all cropped images"""
        if not self.cropped_images:
            messagebox.showwarning("No Images", "No images to save")
            return

        try:
            output_folder = self.settings.get("output_folder", "Scans")
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            date_str = datetime.now().strftime("%Y.%m.%d")
            doc_name = self.doc_name_var.get()
            jpg_quality = self.jpg_quality_var.get()

            # Get EXIF settings
            exif_date = self.exif_date_var.get().strip()
            exif_title = self.exif_title_var.get().strip()

            # Validate EXIF date format if provided
            exif_datetime = None
            if exif_date:
                try:
                    # Parse and validate date
                    parts = exif_date.split(':')
                    if len(parts) != 3:
                        raise ValueError("Invalid date format")
                    year, month, day = int(parts[0]), int(parts[1]), int(parts[2])
                    # Set time to 12:00:00 AM (00:00:00)
                    exif_datetime = f"{year:04d}:{month:02d}:{day:02d} 00:00:00"
                except:
                    messagebox.showerror("Invalid Date",
                                         "EXIF date must be in format YYYY:MM:DD (e.g., 2024:12:25)")
                    return

            # Find next available document number
            doc_counter = 1
            while True:
                test_filename = f"{date_str} {doc_name}{doc_counter}.jpg"
                test_path = os.path.join(output_folder, test_filename)
                if not os.path.exists(test_path):
                    break
                doc_counter += 1

            # Save each image
            saved_files = []
            for idx, data in enumerate(self.cropped_images):
                image = data['image']
                corners = data['corners']

                # Apply transform and crop
                warped = self.four_point_transform(image, corners)
                crop_px = self.crop_pixels_var.get()
                if crop_px > 0:
                    h, w = warped.shape[:2]
                    if h > 2 * crop_px and w > 2 * crop_px:
                        warped = warped[crop_px:h - crop_px, crop_px:w - crop_px]

                # Save
                output_filename = f"{date_str} {doc_name}{doc_counter + idx}.jpg"
                output_path = os.path.join(output_folder, output_filename)
                cv2.imwrite(output_path, warped, [cv2.IMWRITE_JPEG_QUALITY, jpg_quality])
                saved_files.append(output_path)
                self.log(f"Saved: {output_path}")

            # Apply EXIF metadata if specified
            if (exif_datetime or exif_title) and os.path.exists(EXIFTOOL_PATH):
                self.log("Applying EXIF metadata...")
                for filepath in saved_files:
                    try:
                        cmd = [EXIFTOOL_PATH, "-overwrite_original"]

                        if exif_datetime:
                            cmd.extend([f"-DateTimeOriginal={exif_datetime}",
                                        f"-CreateDate={exif_datetime}",
                                        f"-ModifyDate={exif_datetime}"])

                        if exif_title:
                            cmd.extend([f"-Title={exif_title}",
                                        f"-XPTitle={exif_title}"])

                        cmd.append(filepath)

                        result = subprocess.run(cmd, capture_output=True, text=True)
                        if result.returncode == 0:
                            self.log(f"  EXIF applied to {os.path.basename(filepath)}")
                        else:
                            self.log(f"  Warning: EXIF failed for {os.path.basename(filepath)}")
                    except Exception as e:
                        self.log(f"  Warning: Could not apply EXIF to {os.path.basename(filepath)}: {e}")
            elif (exif_datetime or exif_title) and not os.path.exists(EXIFTOOL_PATH):
                self.log(f"Warning: exiftool not found at {EXIFTOOL_PATH}")

            messagebox.showinfo("Success", f"Saved {len(self.cropped_images)} image(s)")
            self.reset_preview()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save images: {str(e)}")

    def rescan(self):
        """Discard current scan and allow rescanning"""
        if messagebox.askyesno("Rescan", "Discard current scan and rescan?"):
            self.reset_preview()
            self.log("Ready to scan again")

    def reset_preview(self):
        """Reset preview state"""
        self.cropped_images = []
        self.current_preview_index = 0
        self.full_scan_image = None
        self.canvas.delete("all")
        self.image_label.config(text="No images")
        self.prev_btn.config(state='disabled')
        self.next_btn.config(state='disabled')
        self.rotate_cw_btn.config(state='disabled')
        self.rotate_ccw_btn.config(state='disabled')
        self.adjust_btn.config(state='disabled')
        self.add_btn.config(state='disabled')
        self.remove_btn.config(state='disabled')
        self.keep_btn.config(state='disabled')
        self.rescan_btn.config(state='disabled')


if __name__ == "__main__":
    root = tk.Tk()
    app = ScannerGUI(root)
    root.mainloop()