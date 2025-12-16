import os
import json
import uuid
import shutil
from core import preview_generator

MODELS_DIR = os.path.join(os.getcwd(), 'models')
os.makedirs(MODELS_DIR, exist_ok=True)

METADATA_FILE = os.path.join(MODELS_DIR, 'metadata.json')

def load_metadata():
    if os.path.exists(METADATA_FILE):
        with open(METADATA_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_metadata(data):
    with open(METADATA_FILE, 'w') as f:
        json.dump(data, f, indent=4)

def add_model(file, image, name):
    model_id = uuid.uuid4().hex
    filename = f"{model_id}.docx"
    
    file_path = os.path.join(MODELS_DIR, filename)
    file.save(file_path)
    
    imagename = None
    if image:
        imagename = f"{model_id}.png"
        image_path = os.path.join(MODELS_DIR, imagename)
        image.save(image_path)
    else:
        # Auto-generate
        imagename = f"{model_id}.png"
        image_path = os.path.join(MODELS_DIR, imagename)
        preview_generator.generate_preview(file_path, image_path)
    
    metadata = load_metadata()
    metadata[model_id] = {
        'id': model_id,
        'name': name,
        'filename': filename,
        'image': imagename
    }
    save_metadata(metadata)
    return model_id

def get_models():
    metadata = load_metadata()
    return list(metadata.values())

def get_model_path(model_id):
    metadata = load_metadata()
    if model_id in metadata:
        return os.path.join(MODELS_DIR, metadata[model_id]['filename'])
    return None

def delete_model(model_id):
    metadata = load_metadata()
    if model_id in metadata:
        info = metadata[model_id]
        try:
            os.remove(os.path.join(MODELS_DIR, info['filename']))
            if info['image']:
                os.remove(os.path.join(MODELS_DIR, info['image']))
        except:
            pass
        del metadata[model_id]
        save_metadata(metadata)
        return True
    return False
