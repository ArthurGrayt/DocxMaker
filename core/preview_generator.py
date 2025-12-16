# Actually using fitz or pdf2image requires external tools like Poppler which user might not have.
# Stick to pure python with Pillow + python-docx extraction.

from docx import Document
from PIL import Image, ImageDraw, ImageFont
import os
import io

def generate_preview(docx_path, output_image_path):
    """
    Generates a simple preview image of the DOCX first page.
    """
    try:
        doc = Document(docx_path)
        
        # A4 Size at 72 DPI approx (595x842)
        WIDTH, HEIGHT = 600, 850
        img = Image.new('RGB', (WIDTH, HEIGHT), 'white')
        draw = ImageDraw.Draw(img)
        
        # Simple Logic: 
        # 1. Try to find images in the header/first paragraph and draw them at top
        # 2. Draw some text abstractly
        
        y_offset = 20
        margin = 40
        
        # --- IMAGES (Logo approximation) ---
        # Docx relationships
        # This is hard to place correctly without a rendering engine.
        # We will iterate through relationships and pick the first image found (likely logo)
        # and place it at the top center.
        
        found_image = None
        for rel in doc.part.rels.values():
             if "image" in rel.target_ref:
                 try:
                     image_data = rel.target_part.blob
                     found_image = Image.open(io.BytesIO(image_data))
                     break
                 except:
                     continue
        
        if found_image:
            # Resize logic
            found_image.thumbnail((200, 100))
            w, h = found_image.size
            # Center at top
            img.paste(found_image, ((WIDTH - w)//2, y_offset))
            y_offset += h + 20
            
        # --- TEXT (Paragraphs) ---
        # Load a default font
        try:
            font = ImageFont.truetype("arial.ttf", 10)
            heading_font = ImageFont.truetype("arial.ttf", 14)
        except:
            font = ImageFont.load_default()
            heading_font = ImageFont.load_default()

        # Read first few paragraphs
        for p in doc.paragraphs[:15]:
            text = p.text.strip()
            if not text:
                continue
            
            # Draw text lines
            # Very naive wrapping
            words = text.split()
            line = ""
            for word in words:
                test_line = line + word + " "
                # Measure (approx)
                if len(test_line) * 6 > (WIDTH - 2*margin):
                    draw.text((margin, y_offset), line, fill="black", font=font)
                    line = word + " "
                    y_offset += 15
                else:
                    line = test_line
            draw.text((margin, y_offset), line, fill="black", font=font)
            y_offset += 15
            
            if y_offset > HEIGHT - 50:
                break

        # Save
        img.save(output_image_path)
        return True
        
    except Exception as e:
        print(f"Error generating preview: {e}")
        # Create a fallback image
        img = Image.new('RGB', (300, 400), '#f0f0f0')
        d = ImageDraw.Draw(img)
        d.text((10, 150), "Preview Inválido", fill="gray")
        img.save(output_image_path)
        return False
