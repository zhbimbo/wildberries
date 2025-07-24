# backend/app.py
import os
import uuid
import logging
from flask import Flask, request, jsonify, send_from_directory, abort
from werkzeug.utils import secure_filename
from flask_cors import CORS  # Импортируем CORS

# Импортируем вашу функцию обработки
try:
    from processor import process_wb_report_file
    PROCESSOR_AVAILABLE = True
except ImportError as e:
    logging.error(f"Не удалось импортировать processor: {e}")
    PROCESSOR_AVAILABLE = False

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# --- Настройка CORS ---
# Разрешаем запросы с любого origin (можно ограничить в production)
CORS(app, origins=["https://zhbimbo.github.io"],  # Указываем ваш домен GitHub Pages
     methods=["GET", "POST", "OPTIONS"], 
     allow_headers=["Content-Type"])

# --- Конфигурация ---
# Используем переменные окружения от Render или значения по умолчанию
UPLOAD_FOLDER = os.environ.get('UPLOAD_FOLDER', '/tmp/uploads')
RESULT_FOLDER = os.environ.get('RESULT_FOLDER', '/tmp/results')
MAX_CONTENT_LENGTH = int(os.environ.get('MAX_CONTENT_LENGTH', 10 * 1024 * 1024)) # 10MB по умолчанию

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

ALLOWED_EXTENSIONS = {'xlsx', 'csv'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET'])
def home():
    return jsonify({"message": "API для обработки отчетов Wildberries запущен!"})

@app.route('/api/upload', methods=['POST', 'OPTIONS'])
def upload_file():
    # Обработка preflight OPTIONS запроса для CORS
    if request.method == 'OPTIONS':
        # Поскольку мы используем flask-cors, он должен автоматически обработать это
        # Но явно вернем 200 OK для уверенности
        return jsonify({"status": "OK"}), 200
        
    logger.info("Получен запрос на загрузку файла")
    
    if not PROCESSOR_AVAILABLE:
        logger.error("Модуль processor недоступен")
        return jsonify({'error': 'Сервис обработки временно недоступен'}), 500

    if 'file' not in request.files:
        logger.warning("Файл не найден в запросе")
        return jsonify({'error': 'Файл не найден в запросе'}), 400

    file = request.files['file']

    if file.filename == '':
        logger.warning("Файл не выбран")
        return jsonify({'error': 'Файл не выбран'}), 400

    if file and allowed_file(file.filename):
        try:
            original_filename = secure_filename(file.filename)
            unique_id = str(uuid.uuid4())
            file_extension = original_filename.rsplit('.', 1)[1].lower()
            unique_filename = f"{unique_id}.{file_extension}"
            
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(file_path)
            logger.info(f"Файл сохранен: {file_path}")

            # Обработка файла
            result_filename = f"результат_{unique_id}.xlsx"
            result_path = os.path.join(app.config['RESULT_FOLDER'], result_filename)
            
            logger.info("Начало обработки файла...")
            process_wb_report_file(file_path, result_path)  # Ваша функция
            logger.info(f"Файл обработан, результат: {result_path}")

            # Очистка временного загруженного файла
            if os.path.exists(file_path):
                os.remove(file_path)
                logger.info("Временный файл удален")

            return jsonify({
                'message': 'Файл успешно обработан',
                'result_filename': result_filename,
                'download_url': f"/api/download/{result_filename}"
            }), 200

        except Exception as e:
            logger.error(f"Ошибка обработки файла: {e}", exc_info=True)
            # Пытаемся удалить временные файлы в случае ошибки
            if 'file_path' in locals() and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    logger.info("Временный файл удален после ошибки")
                except Exception as remove_error:
                    logger.error(f"Ошибка при удалении временного файла: {remove_error}")
            return jsonify({'error': f'Ошибка обработки файла: {str(e)}'}), 500

    else:
        return jsonify({'error': 'Недопустимый тип файла. Разрешены только .xlsx и .csv'}), 400

@app.route('/api/download/<filename>')
def download_file(filename):
    logger.info(f"Запрос на скачивание файла: {filename}")
    try:
        # Защита от path traversal
        safe_path = os.path.join(app.config['RESULT_FOLDER'], os.path.basename(filename))
        if os.path.exists(safe_path) and os.path.isfile(safe_path):
            logger.info(f"Файл найден, отправляем: {safe_path}")
            return send_from_directory(app.config['RESULT_FOLDER'], os.path.basename(filename), as_attachment=True)
        else:
            logger.warning(f"Файл не найден: {safe_path}")
            return jsonify({'error': 'Файл не найден'}), 404
    except Exception as e:
        logger.error(f"Ошибка скачивания файла: {e}", exc_info=True)
        return jsonify({'error': 'Ошибка при скачивании файла'}), 500

# Health check endpoint для Render
@app.route('/healthz')
def health_check():
    return jsonify({"status": "healthy"}), 200

if __name__ == '__main__':
    # Render автоматически устанавливает порт
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
