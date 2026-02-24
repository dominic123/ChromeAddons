"""
Simple script to generate icon files for the Chrome extension.
Requires: pip install Pillow
Run: python generate-icons.py
"""

try:
    from PIL import Image, ImageDraw

    # Create icons at 3 sizes
    sizes = [16, 48, 128]

    for size in sizes:
        # Create a new image with transparent background
        img = Image.new('RGBA', (size, size), (255, 0, 0, 255))
        draw = ImageDraw.Draw(img)

        # Draw red background circle
        margin = size // 16
        draw.ellipse([margin, margin, size - margin, size - margin], fill=(255, 0, 0, 255))

        # Draw white "X" for block symbol
        center = size // 2
        line_width = max(2, size // 8)
        start = size // 4
        end = size * 3 // 4

        # Draw the X
        draw.line([start, start, end, end], fill=(255, 255, 255, 255), width=line_width)
        draw.line([end, start, start, end], fill=(255, 255, 255, 255), width=line_width)

        # Save the icon
        img.save(f'icon{size}.png', 'PNG')
        print(f'Created icon{size}.png')

    print('All icons created successfully!')

except ImportError:
    print('Pillow library not found. Install it with: pip install Pillow')
    print('Or use any image editor to create 16x16, 48x48, and 128x128 PNG files named icon16.png, icon48.png, icon128.png')
