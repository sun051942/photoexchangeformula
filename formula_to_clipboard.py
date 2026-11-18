"""Copy a formula image to the Windows clipboard."""

import io

import win32clipboard
from PIL import Image, ImageOps


def copy_image_to_clipboard(image_path, crop_box=None):
    """Copy an image as DIB, optionally cropping it first."""
    try:
        with Image.open(image_path) as source:
            image = ImageOps.exif_transpose(source).convert("RGBA")
        if crop_box:
            image = image.crop(crop_box)
        background = Image.new("RGBA", image.size, "white")
        image = Image.alpha_composite(background, image).convert("RGB")

        output = io.BytesIO()
        image.save(output, "BMP")
        data = output.getvalue()[14:]  # DIB excludes the 14-byte BMP file header.

        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        finally:
            win32clipboard.CloseClipboard()
        return True
    except Exception as exc:
        print(f"复制到剪贴板失败: {exc}")
        return False
