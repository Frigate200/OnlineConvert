# Используем легкий образ Python на базе Debian
FROM python:3.10-slim-bookworm

LABEL description="Office Document Suite, PDF Converter, Stamping and OCR Service"

# 1. Установка системных зависимостей:
# - poppler-utils: рендеринг PDF
# - tesseract-ocr (rus + eng): распознавание текста
# - default-jre-headless: Java для работы tabula-py
# - libgl1-mesa-glx, libglib2.0-0: библиотеки для OpenCV
# - fonts-dejavu-core, fonts-liberation: кириллические шрифты
RUN apt-get update && apt-get install -y --no-install-recommends \
    poppler-utils \
    tesseract-ocr \
    tesseract-ocr-rus \
    tesseract-ocr-eng \
    default-jre-headless \
    libgl1-mesa-glx \
    libglib2.0-0 \
    fonts-dejavu-core \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# 2. Рабочая директория
WORKDIR /app

# 3. Установка Python зависимостей
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 4. Копирование исходного кода (.dockerignore исключает venv, __pycache__, .git)
COPY . .

# 5. Папка для временных файлов
RUN mkdir -p uploads

EXPOSE 5000

# 6. Запуск через Gunicorn (4 воркера для одновременной работы нескольких сотрудников)
CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:5000", "--timeout", "180", "main:app"]
