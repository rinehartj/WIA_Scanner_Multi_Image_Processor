import win32com.client
from datetime import datetime
import os
import cv2
import numpy as np

# ==================== SCAN SETTINGS ====================
DPI = 800  # Scanning resolution (e.g., 150, 300, 600, 1200)
COLOR_MODE = "Color"  # "Color", "Grayscale", or "BlackAndWhite"
DOCUMENT_NAME = "Document"  # Base name for the scanned file
OUTPUT_FOLDER = "Scans"  # Folder where scans will be saved
# =======================================================

# ==================== CROP SETTINGS ====================
ENABLE_AUTO_CROP = True  # Enable automatic document detection and cropping
JPG_QUALITY = 98  # JPEG compression quality (0-100, higher = better)
MIN_AREA_RATIO = 0.01  # Minimum document size as ratio of image (0.01 = 1%)
CROP_PIXELS = 5  # Pixels to crop from all sides (0 = no additional crop)
# =======================================================

# WIA Constants
WIA_INTENT_NONE = 0x00000000
WIA_INTENT_IMAGE_TYPE_COLOR = 0x00000001
WIA_INTENT_IMAGE_TYPE_GRAYSCALE = 0x00000002
WIA_INTENT_IMAGE_TYPE_TEXT = 0x00000004


def setup_scanner():
    """Initialize WIA and find the scanner."""
    try:
        # Create WIA device manager
        device_manager = win32com.client.Dispatch("WIA.DeviceManager")

        # Find scanner device
        scanner = None
        for i in range(1, device_manager.DeviceInfos.Count + 1):
            device_info = device_manager.DeviceInfos.Item(i)
            if "3200" in device_info.Properties("Name").Value or \
                    "Epson" in device_info.Properties("Name").Value:
                scanner = device_info.Connect()
                print(f"Found scanner: {device_info.Properties('Name').Value}")
                break

        if scanner is None:
            # If specific scanner not found, use first available scanner
            if device_manager.DeviceInfos.Count > 0:
                scanner = device_manager.DeviceInfos.Item(1).Connect()
                print(f"Using scanner: {device_manager.DeviceInfos.Item(1).Properties('Name').Value}")
            else:
                print("No scanner found!")
                return None

        return scanner
    except Exception as e:
        print(f"Error setting up scanner: {e}")
        return None


def list_properties(item):
    """List all available properties for debugging."""
    print("\nAvailable properties:")
    print("-" * 60)
    try:
        for i in range(1, item.Properties.Count + 1):
            prop = item.Properties(i)
            try:
                print(f"ID: {prop.PropertyID}, Name: {prop.Name}, Value: {prop.Value}")
            except:
                print(f"ID: {prop.PropertyID}, Name: {prop.Name}")
    except Exception as e:
        print(f"Error listing properties: {e}")
    print("-" * 60)


def set_property_by_id(item, prop_id, value, prop_name="Unknown"):
    """Safely set a property by ID."""
    try:
        for i in range(1, item.Properties.Count + 1):
            prop = item.Properties(i)
            if prop.PropertyID == prop_id:
                prop.Value = value
                print(f"Set {prop_name} (ID: {prop_id}) to: {value}")
                return True
        print(f"Property {prop_name} (ID: {prop_id}) not found")
        return False
    except Exception as e:
        print(f"Error setting {prop_name} (ID: {prop_id}): {e}")
        return False


def configure_scan_properties(item):
    """Configure scanning properties."""
    print("\nConfiguring scan properties...")

    # Uncomment the next line to see all available properties
    # list_properties(item)

    success = True

    # Set horizontal resolution (Property ID 6147)
    if not set_property_by_id(item, 6147, DPI, "Horizontal Resolution"):
        success = False

    # Set vertical resolution (Property ID 6148)
    if not set_property_by_id(item, 6148, DPI, "Vertical Resolution"):
        success = False

    # Set color mode intent (Property ID 6146)
    intent_value = WIA_INTENT_IMAGE_TYPE_COLOR
    if COLOR_MODE == "Color":
        intent_value = WIA_INTENT_IMAGE_TYPE_COLOR
    elif COLOR_MODE == "Grayscale":
        intent_value = WIA_INTENT_IMAGE_TYPE_GRAYSCALE
    elif COLOR_MODE == "BlackAndWhite":
        intent_value = WIA_INTENT_IMAGE_TYPE_TEXT

    set_property_by_id(item, 6146, intent_value, "Color Mode")

    # Optional: Set scan area starting position and extent
    # Uncomment these if you want to define a specific scan area
    # set_property_by_id(item, 6149, 0, "X Position")
    # set_property_by_id(item, 6150, 0, "Y Position")
    # set_property_by_id(item, 6151, int(8.5 * DPI), "X Extent")  # 8.5 inches
    # set_property_by_id(item, 6152, int(11 * DPI), "Y Extent")   # 11 inches

    if not success:
        print("\nWARNING: Some properties could not be set.")
        print("The scan will proceed with default settings.")
        print("Uncomment 'list_properties(item)' above to see available properties.")


def scan_document(scanner):
    """Perform the scan operation."""
    try:
        # Get the first scanner item (flatbed)
        item = scanner.Items(1)

        # Configure scan properties
        configure_scan_properties(item)

        print("\nStarting scan...")

        # Perform the scan
        image = item.Transfer("{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}")  # TIFF format GUID

        print("Scan completed!")
        return image
    except Exception as e:
        print(f"Error during scan: {e}")
        return None


def save_image(image):
    """Save the scanned image as TIFF."""
    try:
        # Create output folder if it doesn't exist
        if not os.path.exists(OUTPUT_FOLDER):
            os.makedirs(OUTPUT_FOLDER)

        # Generate filename with date
        date_str = datetime.now().strftime("%Y.%m.%d")
        filename = f"{date_str} {DOCUMENT_NAME}.tif"
        filepath = os.path.join(OUTPUT_FOLDER, filename)

        # Check if file exists and add number suffix if needed
        counter = 1
        while os.path.exists(filepath):
            filename = f"{date_str} {DOCUMENT_NAME}_{counter}.tif"
            filepath = os.path.join(OUTPUT_FOLDER, filename)
            counter += 1

        # Save the image
        image.SaveFile(filepath)
        print(f"Image saved to: {filepath}")
        return filepath
    except Exception as e:
        print(f"Error saving image: {e}")
        return None


def order_points(pts):
    """Order points in the order: top-left, top-right, bottom-right, bottom-left."""
    rect = np.zeros((4, 2), dtype="float32")

    # Sum and diff to find corners
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]  # top-left
    rect[2] = pts[np.argmax(s)]  # bottom-right

    diff = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(diff)]  # top-right
    rect[3] = pts[np.argmax(diff)]  # bottom-left

    return rect


def four_point_transform(image, pts):
    """Apply perspective transform to get top-down view of document."""
    rect = order_points(pts)
    (tl, tr, br, bl) = rect

    # Compute width of new image
    widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
    widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
    maxWidth = max(int(widthA), int(widthB))

    # Compute height of new image
    heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
    heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
    maxHeight = max(int(heightA), int(heightB))

    # Construct destination points
    dst = np.array([
        [0, 0],
        [maxWidth - 1, 0],
        [maxWidth - 1, maxHeight - 1],
        [0, maxHeight - 1]
    ], dtype="float32")

    # Compute perspective transform matrix and apply it
    M = cv2.getPerspectiveTransform(rect, dst)
    warped = cv2.warpPerspective(image, M, (maxWidth, maxHeight))

    return warped


def detect_and_crop_documents(image_path):
    """
    Detects one or more documents in a flatbed scanner image.
    Crops and auto-rotates each document, saving as separate JPEG files.
    """
    print("\n" + "=" * 50)
    print("Document Detection and Cropping")
    print("=" * 50)

    # Load image
    image = cv2.imread(image_path)
    if image is None:
        print(f"ERROR: Could not open {image_path}")
        return []

    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # Adaptive thresholding
    blur = cv2.GaussianBlur(gray, (7, 7), 0)
    thresh = cv2.adaptiveThreshold(
        blur, 255,
        cv2.ADAPTIVE_THRESH_MEAN_C,
        cv2.THRESH_BINARY_INV,
        blockSize=51,
        C=10
    )

    # Morphological closing
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (25, 25))
    closed = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)

    # Find contours
    contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    h, w = gray.shape
    img_area = h * w
    min_area = img_area * MIN_AREA_RATIO

    # Collect document boxes
    boxes = []
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area < min_area:
            continue

        rect = cv2.minAreaRect(cnt)
        box = cv2.boxPoints(rect)
        box = np.float32(box)
        boxes.append(box)

    print(f"Detected {len(boxes)} document(s)")

    # Process each detected document
    saved_files = []
    date_str = datetime.now().strftime("%Y.%m.%d")

    # Find the next available document number
    doc_counter = 1
    while True:
        test_filename = f"{date_str} {DOCUMENT_NAME}{doc_counter}.jpg"
        test_path = os.path.join(OUTPUT_FOLDER, test_filename)
        if not os.path.exists(test_path):
            break
        doc_counter += 1

    for idx, box in enumerate(boxes):
        # Extract and deskew the document
        cropped = four_point_transform(image, box)

        # Apply additional pixel crop if specified
        if CROP_PIXELS > 0:
            h, w = cropped.shape[:2]
            if h > 2 * CROP_PIXELS and w > 2 * CROP_PIXELS:
                cropped = cropped[CROP_PIXELS:h - CROP_PIXELS, CROP_PIXELS:w - CROP_PIXELS]
                print(f"  Applied {CROP_PIXELS}px crop to document {idx + 1}")

        # Generate output filename with incrementing document number
        output_filename = f"{date_str} {DOCUMENT_NAME}{doc_counter + idx}.jpg"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        # Save as JPEG with specified quality
        cv2.imwrite(output_path, cropped, [cv2.IMWRITE_JPEG_QUALITY, JPG_QUALITY])
        print(f"  Saved cropped document: {output_path}")
        saved_files.append(output_path)

    return saved_files


def main():
    """Main function to run the scanning process."""
    print("=" * 50)
    print("WIA Scanner with Auto-Crop")
    print("=" * 50)
    print()

    # Setup scanner
    scanner = setup_scanner()
    if scanner is None:
        print("Failed to initialize scanner. Exiting.")
        return

    print()

    # Scan document
    image = scan_document(scanner)
    if image is None:
        print("Scan failed. Exiting.")
        return

    print()

    # Save initial TIFF image
    tiff_filepath = save_image(image)
    if not tiff_filepath:
        print("Failed to save image.")
        return

    print()
    print("=" * 50)
    print("Initial scan completed successfully!")
    print("=" * 50)

    # Auto-crop if enabled
    if ENABLE_AUTO_CROP:
        try:
            cropped_files = detect_and_crop_documents(tiff_filepath)

            if cropped_files:
                print()
                print("=" * 50)
                print("Processing Complete!")
                print("=" * 50)
                print(f"\nOriginal TIFF: {tiff_filepath}")
                print(f"Cropped documents: {len(cropped_files)}")
                for f in cropped_files:
                    print(f"  - {f}")
            else:
                print("\nNo documents detected for cropping.")
                print(f"Original TIFF saved: {tiff_filepath}")
        except Exception as e:
            print(f"\nError during auto-crop: {e}")
            print(f"Original TIFF saved: {tiff_filepath}")
    else:
        print(f"\nAuto-crop disabled. Original TIFF saved: {tiff_filepath}")


if __name__ == "__main__":
    main()