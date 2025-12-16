import os
import uuid
import zipfile
from flask import Flask, render_template, request, send_file, after_this_request
from core.processor import process_docx

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')
app.config['PROCESSED_FOLDER'] = os.path.join(os.getcwd(), 'processed')

# Ensure directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['PROCESSED_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return {'success': False, 'message': 'Nenhum arquivo enviado.'}, 400
    
    files = request.files.getlist('file')
    if not files or files[0].filename == '':
        return {'success': False, 'message': 'Nenhum arquivo selecionado.'}, 400
        
    processed_files = []
    errors = []
    
    # Process each file
    for file in files:
        if file and file.filename.endswith('.docx'):
            try:
                # Keep original filename for the zip but make unique system path
                original_filename = file.filename
                unique_id = uuid.uuid4().hex
                
                input_filename = f"{unique_id}_{original_filename}"
                input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
                
                output_filename = f"unlocked_{original_filename}"
                output_path = os.path.join(app.config['PROCESSED_FOLDER'], f"{unique_id}_{output_filename}")
                
                file.save(input_path)
                
                success, message = process_docx(input_path, output_path)
                
                if success:
                    processed_files.append({
                        'path': output_path,
                        'name': output_filename
                    })
                else:
                    errors.append(f"Erro em {original_filename}: {message}")
            except Exception as e:
                errors.append(f"Falaha ao processar {original_filename}: {str(e)}")
        else:
            errors.append(f"Arquivo ignorado (não é .docx): {file.filename}")

    if not processed_files:
        return {'success': False, 'message': f'Falha ao processar arquivos. Erros: {"; ".join(errors)}'}, 500

    # Create ZIP
    try:
        if len(processed_files) > 1:
            zip_filename = f"unlocked_files_{uuid.uuid4().hex}.zip"
            zip_path = os.path.join(app.config['PROCESSED_FOLDER'], zip_filename)
            
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for p_file in processed_files:
                    zipf.write(p_file['path'], p_file['name'])
            
            download_url = f'/download/{zip_filename}'
        else:
            # Single file, just send it directly (renamed to nice name in download)
            # Actually, to avoid complexity with UUIDs in paths, we create a copy or send the one we have
            # Let's keep the logic simple: verify if we want to force ZIP or not.
            # User asked to "zip all generated docx", implying even for one? Or just for multiple?
            # "o app deverá juntar todos os docx gerados num zip". Plural suggests multiple.
            # But usually for single file, direct download is better.
            # Let's do ZIP always for consistency if that's what user requested?
            # "juntar todos os docx gerados num zip e baixar" -> Let's force zip for consistency with "bundle" request.
            # Actually, standard UX: if 1 file -> download file. If >1 -> download zip.
            # But let's follow explicit "juntar todos... num zip". Maybe safe to always zip.
            # Let's optimize: If > 1 ZIP. If = 1 Direct.
            
            p_file = processed_files[0]
            # Rename specifically for download?
            # The download route takes the filename on disk.
            # Our filename on disk has UUID prefix.
            # We want to serve it with a clean name.
            # Let's use ZIP for everything to handle the "clean name" requirement easily inside the zip.
            
            zip_filename = f"unlocked_files_{uuid.uuid4().hex}.zip"
            zip_path = os.path.join(app.config['PROCESSED_FOLDER'], zip_filename)
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                 for p_file in processed_files:
                    zipf.write(p_file['path'], p_file['name'])
            download_url = f'/download/{zip_filename}'

        msg = "Arquivos processados."
        if errors:
            msg += f" (Alguns arquivos falharam: {len(errors)})"

        return {'success': True, 'download_url': download_url, 'message': msg}
        
    except Exception as e:
         return {'success': False, 'message': f'Erro ao criar arquivo para download: {str(e)}'}, 500

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(app.config['PROCESSED_FOLDER'], filename)
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
