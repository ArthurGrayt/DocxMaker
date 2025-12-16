from docx import Document
from docx.shared import Pt

def process_docx(input_path, output_path):
    """
    Reads a DOCX file, extracts text preserving basic formatting (bold/italic)
    and writes it to a new DOCX file with a clean footer.
    """
    try:
        source_doc = Document(input_path)
        new_doc = Document()
        
        # Iterate over all paragraphs in the source document
        for para in source_doc.paragraphs:
            # Create a new paragraph in the new document
            new_para = new_doc.add_paragraph()
            
            # Copy alignment/style if needed, but for now we focus on run content
            # formatting to ensure clean XML structure.
            new_para.alignment = para.alignment
            new_para.style = para.style.name if para.style.name in new_doc.styles else 'Normal'
            
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                new_run.bold = run.bold
                new_run.italic = run.italic
                new_run.underline = run.underline
                new_run.font.name = run.font.name
                if run.font.size:
                    new_run.font.size = run.font.size
                if run.font.color.rgb:
                    new_run.font.color.rgb = run.font.color.rgb

        # Copy footer from the first section of the source document
        if source_doc.sections and source_doc.sections[0].footer:
            source_footer = source_doc.sections[0].footer
            target_footer = new_doc.sections[0].footer
            target_footer.is_linked_to_previous = False # Ensure independent footer

            # Clear any default empty paragraph if necessary, though newly created doc usually has one empty para
            # We simply append new content or overwrite the existing empty one if needed.
            # But safer to just clear text of the first one if it exists or iterate.
            
            # Let's iterate source footer paragraphs and copy them
            if len(target_footer.paragraphs) > 0:
                target_footer.paragraphs[0].text = "" # Clear default empty paragraph

            for i, para in enumerate(source_footer.paragraphs):
                # If it's the first paragraph, use the existing one (already cleared), else add new
                if i == 0 and len(target_footer.paragraphs) > 0:
                    new_para = target_footer.paragraphs[0]
                else:
                    new_para = target_footer.add_paragraph()
                
                new_para.alignment = para.alignment
                
                for run in para.runs:
                    new_run = new_para.add_run(run.text)
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.underline = run.underline
                    new_run.font.name = run.font.name
                    if run.font.size:
                        new_run.font.size = run.font.size
                    if run.font.color.rgb:
                        new_run.font.color.rgb = run.font.color.rgb

        new_doc.save(output_path)
        return True, "Arquivo processado com sucesso."
        
    except Exception as e:
        return False, str(e)
