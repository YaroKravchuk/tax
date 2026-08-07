"""Turns photographed signatures into transparent-background PNGs with black ink."""
import os
import sys

import numpy as np
from PIL import Image, ImageFilter


def otsu_threshold(values):
    """Finds the brightness that best splits the image into ink and paper."""
    histogram, _ = np.histogram(values, bins=256, range=(0, 256))
    total = values.size
    weighted_total = np.dot(np.arange(256), histogram)

    ink_weight = 0.0
    ink_sum = 0.0
    best_variance = -1.0
    threshold = 0

    for level in range(256):
        ink_weight += histogram[level]
        paper_weight = total - ink_weight
        if ink_weight == 0 or paper_weight == 0:
            continue
        ink_sum += level * histogram[level]
        between = (ink_weight * paper_weight
                   * (ink_sum / ink_weight - (weighted_total - ink_sum) / paper_weight) ** 2)
        if between > best_variance:
            best_variance, threshold = between, level

    return threshold


def clean_signature(src_path):
    gray = Image.open(src_path).convert('L')

    # Estimate the paper behind the ink. A max filter wider than a pen stroke wipes the
    # strokes out, leaving just the lighting across the page; blurring smooths the result.
    background = gray.filter(ImageFilter.MaxFilter(11)).filter(ImageFilter.GaussianBlur(6))

    ink = np.asarray(gray, dtype=float)
    paper = np.maximum(np.asarray(background, dtype=float), 1.0)

    # Divide out the lighting so paper reads as white everywhere, however the photo was lit
    opacity = 255.0 - np.clip(ink / paper * 255.0, 0, 255)

    # Split ink from paper on this image's own contrast, then ramp either side of that
    # split so stroke edges stay smooth instead of turning into jagged pixels
    split = max(otsu_threshold(opacity.astype(np.uint8)), 12)
    alpha = np.clip((opacity - split * 0.55) / (split * 0.85), 0, 1) * 255

    cleaned = np.zeros(ink.shape + (4,), dtype=np.uint8)  # black ink
    cleaned[..., 3] = alpha.astype(np.uint8)
    return Image.fromarray(cleaned, mode='RGBA')


def trim(image, padding=2):
    """Crops away fully transparent margins so the signature fills its cell."""
    box = image.getchannel('A').point(lambda v: 255 if v > 8 else 0).getbbox()
    if box is None:
        return image
    left, top, right, bottom = box
    return image.crop((max(0, left - padding), max(0, top - padding),
                       min(image.width, right + padding), min(image.height, bottom + padding)))


if __name__ == '__main__':
    src_dir, out_dir = sys.argv[1], sys.argv[2]
    os.makedirs(out_dir, exist_ok=True)
    for name in sorted(os.listdir(src_dir)):
        if not name.endswith('_signature.png'):
            continue
        result = trim(clean_signature(os.path.join(src_dir, name)))
        result.save(os.path.join(out_dir, name))
        alpha = np.asarray(result.getchannel('A'))
        print(f'{name:22} {result.width:>3}x{result.height:<3} '
              f'solid ink {100 * (alpha > 200).mean():4.1f}%  '
              f'clear background {100 * (alpha < 16).mean():4.1f}%')
