from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

def copy_paragraph(source_para, target_container):
    """Copies a paragraph and its runs to a target container (doc or cell)."""
    # Create new paragraph in target
    new_para = target_container.add_paragraph()
    
    # Copy paragraph formatting
    new_para.alignment = source_para.alignment
    # Attempt to preserve style if it exists, otherwise default
    if hasattr(source_para, 'style') and source_para.style:
       try:
           new_para.style = source_para.style.name
       except:
           pass # Style might not exist in new doc, ignore

    for run in source_para.runs:
        new_run = new_para.add_run(run.text)
        new_run.bold = run.bold
        new_run.italic = run.italic
        new_run.underline = run.underline
        new_run.font.name = run.font.name
        if run.font.size:
            new_run.font.size = run.font.size
        if run.font.color.rgb:
            new_run.font.color.rgb = run.font.color.rgb

def copy_table(source_table, target_container):
    """Copies a table structure and content."""
    rows = len(source_table.rows)
    cols = len(source_table.columns) if rows > 0 else 0
    
    if rows == 0 or cols == 0:
        return

    new_table = target_container.add_table(rows=0, cols=cols)
    new_table.style = source_table.style.name if hasattr(source_table, 'style') else 'Table Grid'
    
    for src_row in source_table.rows:
        new_row = new_table.add_row()
        for i, src_cell in enumerate(src_row.cells):
            if i < len(new_row.cells):
                tgt_cell = new_row.cells[i]
                # Clear default paragraph in new cell
                # Safe clear of children
                for child in list(tgt_cell._element):
                    tgt_cell._element.remove(child)
                
                # Iterate content of source cell (paragraphs and nested tables)
                for child in src_cell._element.getchildren():
                    if isinstance(child, CT_P):
                        p = Paragraph(child, src_cell)
                        copy_paragraph(p, tgt_cell)
                    elif isinstance(child, CT_Tbl):
                        t = Table(child, src_cell)
                        copy_table(t, tgt_cell)

def process_docx(input_path, output_path):
    """
    Reads a DOCX file, extracts text/tables preserving basic formatting
    and writes it to a new DOCX file with a clean footer.
    """
    try:
        source_doc = Document(input_path)
        new_doc = Document()
        
        # Helper to iterate block items (paragraphs and tables)
        def process_container(source_container, target_container):
            for child in source_container._element.body.iterchildren():
                if isinstance(child, CT_P):
                    p = Paragraph(child, source_container)
                    copy_paragraph(p, target_container)
                elif isinstance(child, CT_Tbl):
                    t = Table(child, source_container)
                    copy_table(t, target_container)

        # Process Body
        process_container(source_doc, new_doc)

        # Process Footer
        # We copy from the first section's footer to the new doc's first section footer
        if source_doc.sections and source_doc.sections[0].footer:
            source_footer = source_doc.sections[0].footer
            target_footer = new_doc.sections[0].footer
            target_footer.is_linked_to_previous = False
            
            # Clear default content in target footer
            for child in list(target_footer._element):
                target_footer._element.remove(child)
            
            # Iterate footer children
            for child in source_footer._element.getchildren():
                if isinstance(child, CT_P):
                    p = Paragraph(child, source_footer)
                    copy_paragraph(p, target_footer)
                elif isinstance(child, CT_Tbl):
                    t = Table(child, source_footer)
                    copy_table(t, target_footer)

        new_doc.save(output_path)
        return True, "Arquivo processado com sucesso."
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return False, str(e)
