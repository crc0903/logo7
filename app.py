import streamlit as st
from PIL import Image, ImageChops
import os
import io
import math
from pptx import Presentation
from pptx.util import Inches

PRELOADED_LOGO_DIR = "preloaded_logos"

# Get all valid preloaded logo filenames
def get_preloaded_logo_filenames():
    if not os.path.exists(PRELOADED_LOGO_DIR):
        os.makedirs(PRELOADED_LOGO_DIR)
    return sorted([
        os.path.splitext(f)[0]
        for f in os.listdir(PRELOADED_LOGO_DIR)
        if f.lower().endswith((".png", ".jpg", ".jpeg", ".webp"))
    ], key=lambda x: x.lower())

# Load only the selected logos
def load_selected_preloaded_logos(selected_names):
    images = []
    for name in selected_names:
        for ext in [".png", ".jpg", ".jpeg", ".webp"]:
            path = os.path.join(PRELOADED_LOGO_DIR, name + ext)
            if os.path.exists(path):
                image = Image.open(path).convert("RGBA")
                images.append((name.lower(), image))
                break
    return images

# Trim white or transparent space around the logo
def trim_whitespace(image):
    bg = Image.new(image.mode, image.size, (255, 255, 255, 0))  # transparent background
    diff = ImageChops.difference(image, bg)
    bbox = diff.getbbox()
    if bbox:
        return image.crop(bbox)
    return image

# Resize logo to fit within a 3:1 box
def resize_to_fill_5x2_box(image, cell_width_px, cell_height_px, buffer_ratio=0.7):
    box_ratio = 3 / 1
    max_box_width = int(cell_width_px * buffer_ratio)
    max_box_height = int(cell_height_px * buffer_ratio)

    if max_box_width / box_ratio <= max_box_height:
        box_width = max_box_width
        box_height = int(max_box_width / box_ratio)
    else:
        box_height = max_box_height
        box_width = int(max_box_height * box_ratio)

    img_w, img_h = image.size
    img_ratio = img_w / img_h

    if img_ratio > (box_width / box_height):
