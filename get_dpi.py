from PIL import Image
from PIL import ExifTags # To interpret EXIF tags if necessary

def get_image_dpi(image_path):
    """
    Fetches the DPI of an image.
    Returns a tuple (horizontal_dpi, vertical_dpi) or (dpi, dpi) if uniform,
    or None if DPI information is not found.
    """
    try:
        img = Image.open(image_path)
        dpi_info = img.info.get('dpi')

        if dpi_info:
            if isinstance(dpi_info, tuple) and len(dpi_info) == 2:
                return dpi_info # (horizontal_dpi, vertical_dpi)
            elif isinstance(dpi_info, (int, float)):
                return (dpi_info, dpi_info) # Uniform DPI

        # If 'dpi' is not directly available, try common JFIF tags (for JPEGs)
        if 'jfif_X_density' in img.info and 'jfif_Y_density' in img.info:
            x_density = img.info['jfif_X_density']
            y_density = img.info['jfif_Y_density']
            unit = img.info.get('jfif_unit')

            if unit == 1:  # Dots Per Inch
                return (x_density, y_density)
            elif unit == 2:  # Dots Per Centimeter
                return (x_density * 2.54, y_density * 2.54)

        # Try to get from EXIF data if available (common in TIFFs, some JPEGs)
        try:
            exif_data = img._getexif()
            if exif_data:
                x_resolution_tag = None
                y_resolution_tag = None
                resolution_unit_tag = None

                # Find the tag IDs for resolution
                for tag_id, name in ExifTags.TAGS.items():
                    if name == "XResolution":
                        x_resolution_tag = tag_id
                    elif name == "YResolution":
                        y_resolution_tag = tag_id
                    elif name == "ResolutionUnit":
                        resolution_unit_tag = tag_id

                x_resolution_val = exif_data.get(x_resolution_tag)
                y_resolution_val = exif_data.get(y_resolution_tag)
                resolution_unit_val = exif_data.get(resolution_unit_tag)

                if x_resolution_val and y_resolution_val:
                    # Resolution value is often a tuple (numerator, denominator)
                    if isinstance(x_resolution_val, tuple):
                        x_res = x_resolution_val[0] / x_resolution_val[1]
                    else:
                        x_res = x_resolution_val
                    
                    if isinstance(y_resolution_val, tuple):
                        y_res = y_resolution_val[0] / y_resolution_val[1]
                    else:
                        y_res = y_resolution_val

                    if resolution_unit_val == 2:  # Inches
                        return (x_res, y_res)
                    elif resolution_unit_val == 3:  # Centimeters
                        return (x_res * 2.54, y_res * 2.54)
                    else: # Unit not specified or unknown, assume DPI if values look like it
                        # Heuristic: if values are reasonable for DPI (e.g., > 50)
                        if x_res > 50 and y_res > 50:
                             return (x_res, y_res)

        except (AttributeError, TypeError, IndexError, KeyError):
            # Issue getting or parsing EXIF
            pass
            
        return None # DPI not found

    except FileNotFoundError:
        print(f"Error: Image file not found at {image_path}")
        return None
    except Exception as e:
        print(f"An error occurred while processing {image_path}: {e}")
        return None
    finally:
        if 'img' in locals() and img:
            img.close()

# --- How to use the function ---
if __name__ == "__main__":
    # Replace "path/to/your/image_c105a7.png" with the actual path to your image
    image_file_path = "image.png"  # <--- CHANGE THIS

    dpi_values = get_image_dpi(image_file_path)

    if dpi_values:
        print(f"Image: {image_file_path}")
        print(f"Horizontal DPI: {dpi_values[0]}")
        print(f"Vertical DPI: {dpi_values[1]}")
    else:
        print(f"Could not determine DPI for image: {image_file_path}")
        print("The image may not have DPI metadata, or it's stored in an unrecognized format.")