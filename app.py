import os
import uuid
import zipfile
from flask import Flask, render_template, request, send_file, after_this_request
from core.processor import process_docx, process_with_model
from core import model_manager

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')
app.config['PROCESSED_FOLDER'] = os.path.join(os.getcwd(), 'processed')
app.config['MODELS_DIR'] = model_manager.MODELS_DIR

# Ensure directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['PROCESSED_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/models', methods=['GET', 'POST'])
def handle_models():
    if request.method == 'POST':
        if 'file' not in request.files:
            return {'success': False, 'message': 'Arquivo DOCX é obrigatório.'}, 400
        
        file = request.files['file']
        image = request.files.get('image') # Optional now
        name = request.form.get('name', 'Modelo sem nome')
        
        if file.filename == '':
             return {'success': False, 'message': 'Arquivo inválido.'}, 400
             
        if not file.filename.endswith('.docx'):
             return {'success': False, 'message': 'O modelo deve ser .docx'}, 400
             
        # Image logic handled inside manager (auto-generate if None)
        model_id = model_manager.add_model(file, image, name)
        return {'success': True, 'model': {'id': model_id, 'name': name}}

    # GET
    return {'models': model_manager.get_models()}

@app.route('/models/<model_id>/image')
def model_image(model_id):
    models = model_manager.get_models()
    model = next((m for m in models if m['id'] == model_id), None)
    if model and model['image']:
        return send_file(os.path.join(app.config['MODELS_DIR'], model['image']))
    return '', 404

@app.route('/models/<model_id>', methods=['DELETE'])
def delete_model(model_id):
    if model_manager.delete_model(model_id):
         return {'success': True}
    return {'success': False, 'message': 'Modelo não encontrado'}, 404

@app.route('/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return {'success': False, 'message': 'Nenhum arquivo enviado.'}, 400
    
    files = request.files.getlist('file')
    if not files or files[0].filename == '':
        return {'success': False, 'message': 'Nenhum arquivo selecionado.'}, 400
        
    model_id = request.form.get('model_id')
    model_path = None
    if model_id:
        model_path = model_manager.get_model_path(model_id)
        if not model_path:
             return {'success': False, 'message': 'Modelo selecionado não encontrado.'}, 400

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
                
                output_filename = f"processed_{original_filename}" # nicer name
                output_path = os.path.join(app.config['PROCESSED_FOLDER'], f"{unique_id}_{output_filename}")
                
                file.save(input_path)
                
                if model_path:
                    success, message = process_with_model(input_path, model_path, output_path)
                else:
                    success, message = process_docx(input_path, output_path)
                
                if success:
                    processed_files.append({
                        'path': output_path,
                        'name': output_filename
                    })
                else:
                    errors.append(f"Erro em {original_filename}: {message}")
            except Exception as e:
                import traceback
                traceback.print_exc()
                errors.append(f"Falha ao processar {original_filename}: {str(e)}")
        else:
            errors.append(f"Arquivo ignorado (não é .docx): {file.filename}")

    if not processed_files:
        return {'success': False, 'message': f'Falha ao processar arquivos. Erros: {"; ".join(errors)}'}, 500

    # Create ZIP
    try:
        # Always zip
        zip_filename = f"arquivos_processados_{uuid.uuid4().hex}.zip"
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
