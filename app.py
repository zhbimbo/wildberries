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
# Разрешаем запросы с вашего домена GitHub Pages
# Можно указать список origins, если нужно
CORS(app, origins=["https://zhbimbo.github.io", "http://localhost:5173"],  # Добавил localhost для локальной разработки
     methods=["GET", "POST", "OPTIONS"], 
     allow_headers=["Content-Type"])

# --- Конфигурация ---
# Используем переменные окружения от Render или значения по умолчанию
# Для Render лучше использовать /tmp, так как это стандартная директория для временных файлов
UPLOAD_FOLDER = os.environ.get('UPLOAD_FOLDER', '/tmp/uploads')
RESULT_FOLDER = os.environ.get('RESULT_FOLDER', '/tmp/results')
MAX_CONTENT_LENGTH = int(os.environ.get('MAX_CONTENT_LENGTH', 16 * 1024 * 1024)) # 16MB по умолчанию

# Создаем директории, если они не существуют
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

ALLOWED_EXTENSIONS = {'xlsx', 'csv'}

def allowed_file(filename):
    """Проверяет, разрешено ли расширение файла."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET'])
def home():
    """Корневой эндпоинт для проверки работы API."""
    return jsonify({"message": "API для обработки отчетов Wildberries запущен!"})

@app.route('/healthz', methods=['GET'])
def health_check():
    """Health check endpoint для Render."""
    return jsonify({"status": "healthy"}), 200

@app.route('/api/upload', methods=['POST', 'OPTIONS'])
def upload_file():
    """Эндпоинт для загрузки и обработки файла."""
    # Обработка preflight OPTIONS запроса для CORS
    if request.method == 'OPTIONS':
        # Flask-CORS должен обработать это автоматически, но явно вернем 200 OK
        return jsonify({"status": "OK"}), 200
        
    logger.info("Получен запрос на загрузку файла")
    
    if not PROCESSOR_AVAILABLE:
        logger.error("Модуль processor недоступен")
        return jsonify({'error': 'Сервис обработки временно недоступен'}), 500

    # Проверка наличия файла в запросе
    if 'file' not in request.files:
        logger.warning("Файл не найден в запросе")
        return jsonify({'error': 'Файл не найден в запросе'}), 400

    file = request.files['file']

    # Проверка, был ли выбран файл
    if file.filename == '':
        logger.warning("Файл не выбран")
        return jsonify({'error': 'Файл не выбран'}), 400

    # Проверка допустимого типа файла
    if file and allowed_file(file.filename):
        try:
            # Генерируем уникальное имя для временного файла
            # Используем оригинальное расширение
            original_extension = file.filename.rsplit('.', 1)[1].lower()
            unique_id = str(uuid.uuid4())
            temp_filename = f"temp_{unique_id}.{original_extension}"
            temp_file_path = os.path.join(app.config['UPLOAD_FOLDER'], temp_filename)
            
            # Сохраняем загруженный файл
            file.save(temp_file_path)
            logger.info(f"Файл сохранен как временный: {temp_file_path}")

            # Определяем имя файла результата
            result_filename = f"результат_{unique_id}.xlsx"
            result_file_path = os.path.join(app.config['RESULT_FOLDER'], result_filename)
            
            logger.info("Начало обработки файла...")
            # Вызываем вашу функцию обработки
            # Передаем путь к временному файлу и путь для сохранения результата
            process_wb_report_file(temp_file_path, result_file_path)
            logger.info(f"Файл обработан, результат сохранен: {result_file_path}")

            # Очистка временного загруженного файла
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
                logger.info("Временный файл удален")

            # Возвращаем успех с информацией о результате
            return jsonify({
                'message': 'Файл успешно обработан',
                'result_filename': result_filename,
                'download_url': f"/api/download/{result_filename}"
            }), 200

        except Exception as e:
            logger.error(f"Ошибка обработки файла: {e}", exc_info=True)
            # Пытаемся удалить временные файлы в случае ошибки
            if 'temp_file_path' in locals() and os.path.exists(temp_file_path):
                try:
                    os.remove(temp_file_path)
                    logger.info("Временный файл удален после ошибки")
                except Exception as remove_error:
                    logger.error(f"Ошибка при удалении временного файла: {remove_error}")
            return jsonify({'error': f'Ошибка обработки файла: {str(e)}'}), 500

    else:
        return jsonify({'error': 'Недопустимый тип файла. Разрешены только .xlsx и .csv'}), 400

@app.route('/api/download/<filename>')
def download_file(filename):
    """Эндпоинт для скачивания результата обработки."""
    logger.info(f"Запрос на скачивание файла: {filename}")
    try:
        # Защита от path traversal - используем только базовое имя файла
        safe_filename = os.path.basename(filename)
        file_path = os.path.join(app.config['RESULT_FOLDER'], safe_filename)
        
        # Проверяем, существует ли файл и является ли он файлом (а не директорией)
        if os.path.exists(file_path) and os.path.isfile(file_path):
            logger.info(f"Файл найден, отправляем: {file_path}")
            # Отправляем файл как вложение
            return send_from_directory(app.config['RESULT_FOLDER'], safe_filename, as_attachment=True)
        else:
            logger.warning(f"Файл не найден или не является файлом: {file_path}")
            return jsonify({'error': 'Файл не найден'}), 404
    except Exception as e:
        logger.error(f"Ошибка скачивания файла: {e}", exc_info=True)
        return jsonify({'error': 'Ошибка при скачивании файла'}), 500

# --- Для разработки: раздача статики из React build ---
# (Позже фронтенд будет на отдельном порту или сервере)
# FRONTEND_BUILD_DIR = '../frontend/build' # Путь к сборке React
# @app.route('/')
# def index():
#     """Обслуживает React-приложение."""
#     try:
#         return send_from_directory(FRONTEND_BUILD_DIR, 'index.html')
#     except FileNotFoundError:
#         return "Фронтенд не найден. Запустите 'npm run build' в папке frontend.", 404

# @app.route('/<path:filename>')
# def serve_static_files(filename):
#     """Обслуживает статические файлы React."""
#     try:
#         return send_from_directory(FRONTEND_BUILD_DIR, filename)
#     except FileNotFoundError:
#         # Если файл не найден, возвращаем index.html для клиентского роутинга
#         return send_from_directory(FRONTEND_BUILD_DIR, 'index.html')

if __name__ == '__main__':
    # Render автоматически устанавливает порт через переменную окружения PORT
    # По умолчанию используем 5000 для локальной разработки
    port = int(os.environ.get('PORT', 5000))
    # Важно: host='0.0.0.0' позволяет получить доступ к приложению снаружи контейнера
    app.run(host='0.0.0.0', port=port, debug=False) # Установите debug=True для разработки, но False для продакшена
