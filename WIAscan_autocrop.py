import win32com.client
import pythoncom
from datetime import datetime
import os
import cv2
import numpy as np
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from PIL import Image, ImageTk
import json
import threading

# Settings file for persistence
SETTINGS_FILE = "scanner_settings.json"

# WIA Constants
WIA_INTENT_NONE = 0x00000000
WIA_INTENT_IMAGE_TYPE_COLOR = 0x00000001
WIA_INTENT_IMAGE_TYPE_GRAYSCALE = 0x00000002
WIA_INTENT_IMAGE_TYPE_TEXT = 0x00000004


class ScannerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Scanner Controller")
        self.root.geometry("900x700")

        # Scanner variables
        self.scanner = None
        self.tiff_filepath = None
        self.cropped_images = []
        self.current_preview_index = 0
        self.corner_points = []
        self.adjustment_mode = False
        self.selected_corner = None

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
            "output_folder": "Scans"
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

        # Settings Frame
        settings_frame = ttk.LabelFrame(main_frame, text="Scan Settings", padding="10")
        settings_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        # DPI
        ttk.Label(settings_frame, text="DPI:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.dpi_var = tk.IntVar(value=self.settings.get("dpi", 800))
        dpi_spin = ttk.Spinbox(settings_frame, from_=150, to=2400, increment=50,
                               textvariable=self.dpi_var, width=10)
        dpi_spin.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        # Scan Distance
        ttk.Label(settings_frame, text="Scan Distance (inches):").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        self.distance_var = tk.DoubleVar(value=self.settings.get("scan_distance", 3.0))
        distance_spin = ttk.Spinbox(settings_frame, from_=1.0, to=12.0, increment=0.5,
                                    textvariable=self.distance_var, width=10)
        distance_spin.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)

        # JPG Quality
        ttk.Label(settings_frame, text="JPG Quality (%):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.jpg_quality_var = tk.IntVar(value=self.settings.get("jpg_quality", 98))
        jpg_spin = ttk.Spinbox(settings_frame, from_=50, to=100, increment=1,
                               textvariable=self.jpg_quality_var, width=10)
        jpg_spin.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        # Crop Pixels
        ttk.Label(settings_frame, text="Crop Pixels:").grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
        self.crop_pixels_var = tk.IntVar(value=self.settings.get("crop_pixels", 5))
        crop_spin = ttk.Spinbox(settings_frame, from_=0, to=50, increment=1,
                                textvariable=self.crop_pixels_var, width=10)
        crop_spin.grid(row=1, column=3, sticky=tk.W, padx=5, pady=5)

        # Save TIFF checkbox
        self.save_tiff_var = tk.BooleanVar(value=self.settings.get("save_tiff", True))
        tiff_check = ttk.Checkbutton(settings_frame, text="Save TIFF file",
                                     variable=self.save_tiff_var)
        tiff_check.grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        # Document Name
        ttk.Label(settings_frame, text="Document Name:").grid(row=2, column=2, sticky=tk.W, padx=5, pady=5)
        self.doc_name_var = tk.StringVar(value=self.settings.get("document_name", "Document"))
        doc_entry = ttk.Entry(settings_frame, textvariable=self.doc_name_var, width=15)
        doc_entry.grid(row=2, column=3, sticky=tk.W, padx=5, pady=5)

        # Scan Button
        scan_btn = ttk.Button(main_frame, text="Scan", command=self.start_scan, width=20)
        scan_btn.grid(row=1, column=0, columnspan=2, pady=10)

        # Progress/Log Frame
        log_frame = ttk.LabelFrame(main_frame, text="Status", padding="10")
        log_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        main_frame.rowconfigure(2, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, width=80, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Preview Frame
        preview_frame = ttk.LabelFrame(main_frame, text="Preview", padding="10")
        preview_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        main_frame.rowconfigure(3, weight=2)

        # Canvas for image display
        self.canvas = tk.Canvas(preview_frame, bg='gray', width=800, height=300)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Button-1>", self.on_canvas_click)

        # Navigation and action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)

        self.prev_btn = ttk.Button(button_frame, text="← Previous", command=self.prev_image, state='disabled')
        self.prev_btn.pack(side=tk.LEFT, padx=5)

        self.image_label = ttk.Label(button_frame, text="No images")
        self.image_label.pack(side=tk.LEFT, padx=20)

        self.next_btn = ttk.Button(button_frame, text="Next →", command=self.next_image, state='disabled')
        self.next_btn.pack(side=tk.LEFT, padx=5)

        self.adjust_btn = ttk.Button(button_frame, text="Adjust Corners",
                                     command=self.toggle_adjustment_mode, state='disabled')
        self.adjust_btn.pack(side=tk.LEFT, padx=20)

        self.keep_btn = ttk.Button(button_frame, text="Keep All",
                                   command=self.keep_images, state='disabled')
        self.keep_btn.pack(side=tk.LEFT, padx=5)

        self.rescan_btn = ttk.Button(button_frame, text="Rescan",
                                     command=self.rescan, state='disabled')
        self.rescan_btn.pack(side=tk.LEFT, padx=5)

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
                # Save temporary TIFF for processing
                self.log("Creating temporary file for processing...")
                temp_folder = self.settings.get("output_folder", "Scans")
                if not os.path.exists(temp_folder):
                    os.makedirs(temp_folder)
                self.tiff_filepath = os.path.join(temp_folder, "temp_scan.tif")
                image.SaveFile(self.tiff_filepath)

            # Detect and crop documents
            self.log("Detecting documents...")
            self.cropped_images = self.detect_and_crop_documents(self.tiff_filepath)

            if self.cropped_images:
                self.log(f"Detected {len(self.cropped_images)} document(s)")
                self.current_preview_index = 0
                self.root.after(0, self.display_preview)
                self.root.after(0, self.enable_preview_controls)
            else:
                self.log("No documents detected")

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

            cropped_data.append({
                'image': image.copy(),
                'corners': box,
                'original_corners': box.copy()
            })

        return cropped_data

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

        # Draw corner points if in adjustment mode
        if self.adjustment_mode:
            self.draw_corner_points()

        # Update label
        self.image_label.config(text=f"Image {self.current_preview_index + 1} of {len(self.cropped_images)}")

    def draw_corner_points(self):
        """Draw adjustable corner points on original image"""
        if not self.cropped_images:
            return

        data = self.cropped_images[self.current_preview_index]
        image = data['image'].copy()
        corners = data['corners']

        # Draw corners on image
        for i, pt in enumerate(corners):
            cv2.circle(image, tuple(pt.astype(int)), 15, (0, 255, 0), -1)
            cv2.putText(image, str(i + 1), tuple(pt.astype(int)),
                        cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 2)

        # Draw lines between corners
        for i in range(4):
            pt1 = tuple(corners[i].astype(int))
            pt2 = tuple(corners[(i + 1) % 4].astype(int))
            cv2.line(image, pt1, pt2, (0, 255, 0), 3)

        # Convert and display
        display_img = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)

        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        h, w = display_img.shape[:2]

        scale = min(canvas_width / w, canvas_height / h, 1.0)
        self.scale = scale
        new_w = int(w * scale)
        new_h = int(h * scale)

        display_img = cv2.resize(display_img, (new_w, new_h))
        img_pil = Image.fromarray(display_img)
        self.photo = ImageTk.PhotoImage(img_pil)

        self.canvas.delete("all")
        self.canvas.create_image(canvas_width // 2, canvas_height // 2, image=self.photo)

    def on_canvas_click(self, event):
        """Handle canvas click for corner adjustment"""
        if not self.adjustment_mode or not self.cropped_images:
            return

        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        data = self.cropped_images[self.current_preview_index]
        h, w = data['image'].shape[:2]

        scale = min(canvas_width / w, canvas_height / h, 1.0)

        # Convert canvas coordinates to image coordinates
        offset_x = (canvas_width - w * scale) // 2
        offset_y = (canvas_height - h * scale) // 2

        img_x = (event.x - offset_x) / scale
        img_y = (event.y - offset_y) / scale

        # Find nearest corner
        corners = data['corners']
        min_dist = float('inf')
        closest_corner = None

        for i, pt in enumerate(corners):
            dist = np.sqrt((pt[0] - img_x) ** 2 + (pt[1] - img_y) ** 2)
            if dist < min_dist:
                min_dist = dist
                closest_corner = i

        # Update corner position if close enough
        if min_dist < 50 / scale:
            corners[closest_corner] = [img_x, img_y]
            self.draw_corner_points()

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
        if len(self.cropped_images) > 1:
            self.prev_btn.config(state='normal')
            self.next_btn.config(state='normal')

        self.adjust_btn.config(state='normal')
        self.keep_btn.config(state='normal')
        self.rescan_btn.config(state='normal')

    def prev_image(self):
        """Show previous image"""
        if self.current_preview_index > 0:
            self.current_preview_index -= 1
            self.adjustment_mode = False
            self.adjust_btn.config(text="Adjust Corners")
            self.display_preview()

    def next_image(self):
        """Show next image"""
        if self.current_preview_index < len(self.cropped_images) - 1:
            self.current_preview_index += 1
            self.adjustment_mode = False
            self.adjust_btn.config(text="Adjust Corners")
            self.display_preview()

    def toggle_adjustment_mode(self):
        """Toggle corner adjustment mode"""
        self.adjustment_mode = not self.adjustment_mode
        if self.adjustment_mode:
            self.adjust_btn.config(text="Apply Adjustment")
            self.draw_corner_points()
        else:
            self.adjust_btn.config(text="Adjust Corners")
            self.display_preview()

    def keep_images(self):
        """Save all cropped images"""
        if not self.cropped_images:
            return

        try:
            output_folder = self.settings.get("output_folder", "Scans")
            date_str = datetime.now().strftime("%Y.%m.%d")
            doc_name = self.doc_name_var.get()
            jpg_quality = self.jpg_quality_var.get()

            # Find next available document number
            doc_counter = 1
            while True:
                test_filename = f"{date_str} {doc_name}{doc_counter}.jpg"
                test_path = os.path.join(output_folder, test_filename)
                if not os.path.exists(test_path):
                    break
                doc_counter += 1

            # Save each image
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
                self.log(f"Saved: {output_path}")

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
        self.adjustment_mode = False
        self.canvas.delete("all")
        self.image_label.config(text="No images")
        self.prev_btn.config(state='disabled')
        self.next_btn.config(state='disabled')
        self.adjust_btn.config(state='disabled')
        self.keep_btn.config(state='disabled')
        self.rescan_btn.config(state='disabled')


if __name__ == "__main__":
    root = tk.Tk()
    app = ScannerGUI(root)
    root.mainloop()