import colorsys
import configparser
import os
from collections import OrderedDict
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import win32api
import win32com.client
import win32con
import win32gui
import win32ui
from PIL import Image


def get_average_hue(image_data: BytesIO) -> float:
    """
    Calculate the average hue of an image stored in a BytesIO object.

    Args:
        image_data (BytesIO): A BytesIO object containing image data.

    Returns:
        float: The average hue value of the image. Hue is a value from 0.0 to 1.0 representing colors on the color
        wheel.
    """
    with Image.open(image_data) as img:
        if img.mode != "RGB":
            img = img.convert("RGB")

        img = img.resize((1, 1))  # Resize to 1x1 for fast average color calculation
        avg_color = img.getpixel((0, 0))

        return colorsys.rgb_to_hsv(*avg_color)[0]


def sort_icons_by_color(icons: Dict[str, BytesIO]) -> OrderedDict:
    """
    Sorts a dictionary of icons by their average color hue.

    Args:
        icons (Dict[str, BytesIO]): A dictionary where keys are file paths and values are BytesIO objects of the icon
        images.

    Returns:
        OrderedDict: An ordered dictionary of icons sorted by their average hue value.
    """
    return OrderedDict(sorted(icons.items(), key=lambda item: get_average_hue(item[1])))


def get_icon_handle(exe_path: str) -> Optional[int]:
    """
    Extracts the handle of the first large icon from the specified executable file.

    Args:
        exe_path (str): The file path of the executable.

    Returns:
        Optional[int]: A handle to the first large icon found. Returns None if no icon is found.
    """
    large_icons, small_icons = win32gui.ExtractIconEx(exe_path, 0)
    return large_icons[0] if large_icons else (small_icons[0] if small_icons else None)


def create_icon_bitmap(icon_handle: int) -> Tuple[int, int, int]:
    """
    Creates a bitmap from an icon handle.

    Args:
        icon_handle (int): A handle to an icon.

    Returns:
        Tuple[int, int, int]: A tuple containing the device context handle, bitmap handle, and original device context
        handle.
    """
    # Get icon dimensions
    icon_width = win32api.GetSystemMetrics(win32con.SM_CXICON)  # noqa
    icon_height = win32api.GetSystemMetrics(win32con.SM_CYICON)  # noqa

    # Create device context and bitmap for drawing the icon
    dc_handle = win32gui.GetDC(0)
    mem_dc = win32gui.CreateCompatibleDC(dc_handle)
    bitmap = win32gui.CreateCompatibleBitmap(dc_handle, icon_width, icon_height)
    win32gui.SelectObject(mem_dc, bitmap)
    win32gui.DrawIconEx(
        mem_dc, 0, 0, icon_handle, icon_width, icon_height, 0, None, win32con.DI_NORMAL
    )

    return mem_dc, bitmap, dc_handle


def convert_bitmap_to_image(bitmap: int, icon_size: Tuple[int, int]) -> Image:
    """
    Converts a bitmap handle to a PIL Image object.

    Args:
        bitmap (int): A handle to a bitmap.
        icon_size (Tuple[int, int]): The width and height of the icon.

    Returns:
        Image: A PIL Image object created from the bitmap.
    """
    # Convert the bitmap to a string and then to a PIL Image
    bmpstr = win32ui.CreateBitmapFromHandle(bitmap).GetBitmapBits(True)
    return Image.frombuffer("RGBA", icon_size, bmpstr, "raw", "BGRA")


def extract_icon_from_file(exe_path: str) -> Optional[BytesIO]:
    """
    Extracts the largest icon from the given executable file and converts it to PNG format stored in a BytesIO object.

    Args:
        exe_path (str): The file path of the executable.

    Returns:
        Optional[BytesIO]: A BytesIO object containing the PNG data of the icon, or None if no icon is found.
    """
    icon_handle = get_icon_handle(exe_path)
    if not icon_handle:
        return None

    mem_dc, bitmap, dc_handle = create_icon_bitmap(icon_handle)
    icon_width = win32api.GetSystemMetrics(win32con.SM_CXICON)  # noqa
    icon_height = win32api.GetSystemMetrics(win32con.SM_CYICON)  # noqa

    image = convert_bitmap_to_image(bitmap, (icon_width, icon_height))

    # Clean up the resources used for drawing the icon
    win32gui.DeleteObject(bitmap)
    win32gui.DeleteDC(mem_dc)
    win32gui.ReleaseDC(0, dc_handle)
    win32gui.DestroyIcon(icon_handle)

    # Convert the image to PNG and store in BytesIO for later processing
    png_data = BytesIO()
    image.save(png_data, "PNG")
    png_data.seek(0)
    return png_data


def extract_icon_from_shortcut(shortcut_path: str) -> Optional[BytesIO]:
    """
    Extracts the icon from a shortcut file.

    Args:
        shortcut_path (str): The file path of the shortcut (.lnk) file.

    Returns:
        Optional[BytesIO]: A BytesIO object containing the PNG data of the icon.
    """
    shortcut = win32com.client.Dispatch("WScript.Shell").CreateShortCut(shortcut_path)
    icon_location = shortcut.IconLocation
    icon_path, _, icon_index = icon_location.partition(",")

    # If the icon path is empty, default to the shortcut target.
    if not icon_path:
        icon_path = shortcut.TargetPath

    # If there's a specific icon index, use it; otherwise, default to 0.
    icon_index = int(icon_index) if icon_index.isdigit() else 0

    # Check if the icon file exists.
    if os.path.exists(icon_path):
        return extract_icon_from_file(icon_path)

    return None


def extract_icon_from_url(url_path: str) -> Optional[BytesIO]:
    """
    Extracts an icon from a .url file.

    Args:
        url_path (str): The file path of the .url file.

    Returns:
        Optional[BytesIO]: A BytesIO object containing the PNG data of the icon, or None if no icon is found.
    """
    # Parse the .url file to extract the icon file path
    parser = configparser.ConfigParser()
    parser.read(url_path)

    if "InternetShortcut" in parser and "IconFile" in parser["InternetShortcut"]:
        icon_path = parser["InternetShortcut"]["IconFile"]
        if os.path.exists(icon_path):
            return extract_icon_from_file(icon_path)

    return None


def get_icons() -> Tuple[Dict[str, BytesIO], List[str]]:
    """
    Extracts icons from shortcuts on both the user's and the public desktop, and sorts them by color.

    Returns:
        Tuple[Dict[str, BytesIO], List[str]]: A tuple containing a dictionary of sortable icons and a list of
        unsortable file paths.
    """
    icons = {}
    unsortable = []

    # Paths for the user's desktop and the public desktop
    user_desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
    public_desktop = os.path.join(os.environ["PUBLIC"], "Desktop")

    # Combine the files from both desktops
    files = os.listdir(user_desktop) + os.listdir(public_desktop)

    for file in files:
        for desktop in [user_desktop, public_desktop]:
            path = os.path.join(desktop, file)

            if not os.path.exists(path):
                continue

            try:
                if file.endswith(".lnk"):
                    # Use the new function to extract the icon from the shortcut
                    img_bytes = extract_icon_from_shortcut(path)
                    if img_bytes:
                        icons[path] = img_bytes
                elif file.endswith(".url"):
                    # Extract icon from internet shortcut (.url) file
                    img_bytes = extract_icon_from_url(path)
                    if img_bytes:
                        icons[path] = img_bytes
                else:
                    # File type not supported
                    if not path.endswith("desktop.ini"):
                        unsortable.append(path)
            except Exception as e:
                print(f"Error processing {path}: {e}")

    return icons, unsortable


if __name__ == "__main__":
    sortable_icons, unsortable_icons = get_icons()
    sorted_icons = sort_icons_by_color(sortable_icons)

    print("Icons sorted by colour:")
    for i, k in enumerate(sorted_icons.keys(), start=1):
        print(f"{i}. {k.split('\\')[-1].split('.')[0]}")

    if unsortable_icons:
        print("\nUnable to sort (may not have an icon):")
        print(", ".join(u.split("\\")[-1] for u in unsortable_icons))
