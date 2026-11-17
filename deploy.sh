#!/bin/bash
set -e

# Скрипт автоматического развертывания ОфисКонвертер на удаленный сервер
# Использование: ./deploy.sh [user@host] [remote_path] [port]
# Пример: ./deploy.sh root@123.45.67.89 /opt/convert 22

TARGET="${1}"
REMOTE_PATH="${2:-/opt/convert}"
SSH_PORT="${3:-22}"

if [ -z "$TARGET" ]; then
    echo "Использование: ./deploy.sh user@host [remote_path] [port]"
    echo "Пример: ./deploy.sh root@192.168.1.100 /opt/convert 22"
    exit 1
fi

echo "🚀 [1/4] Проверка подключения к $TARGET (порт $SSH_PORT)..."
ssh -p "$SSH_PORT" -o BatchMode=yes -o ConnectTimeout=8 "$TARGET" "echo 'Подключение успешно: \$(uname -a)'" || {
    echo "⚠️ Ошибка подключения по SSH-ключу к $TARGET."
    echo "Убедитесь, что настроен SSH-ключ: ssh-copy-id -p $SSH_PORT $TARGET"
    exit 1
}

echo "📦 [2/4] Создание архива проекта (без локального venv)..."
ARCHIVE_NAME="convert_deploy.tar.gz"
tar -czf "$ARCHIVE_NAME" \
    --exclude='venv' \
    --exclude='.idea' \
    --exclude='__pycache__' \
    --exclude='.git' \
    --exclude='uploads/*' \
    --exclude="$ARCHIVE_NAME" \
    .

echo "📤 [3/4] Отправка файлов на сервер в $REMOTE_PATH..."
ssh -p "$SSH_PORT" "$TARGET" "mkdir -p $REMOTE_PATH"
scp -P "$SSH_PORT" "$ARCHIVE_NAME" "$TARGET:$REMOTE_PATH/"
rm -f "$ARCHIVE_NAME"

echo "⚙️ [4/4] Развертывание и запуск на сервере..."
ssh -p "$SSH_PORT" "$TARGET" bash -s <<EOF
set -e
cd "$REMOTE_PATH"
tar -xzf "$ARCHIVE_NAME"
rm -f "$ARCHIVE_NAME"
mkdir -p uploads

# Проверяем наличие Docker
if command -v docker >/dev/null 2>&1; then
    echo "🐳 Обнаружен Docker. Запуск через Docker Compose..."
    if docker compose version >/dev/null 2>&1; then
        docker compose up -d --build
    else
        docker-compose up -d --build
    fi
    echo "✅ Контейнер успешно запущен!"
    docker ps | grep -E "office-converter|convert" || true
else
    echo "⚠️ Docker не найден. Устанавливаем системные пакеты и настраиваем Python venv..."
    if command -v apt-get >/dev/null 2>&1; then
        apt-get update
        apt-get install -y python3 python3-pip python3-venv poppler-utils tesseract-ocr tesseract-ocr-rus fonts-dejavu-core
    fi
    python3 -m venv venv
    venv/bin/pip install --upgrade pip
    venv/bin/pip install -r requirements.txt
    
    # Запуск Gunicorn в фоне
    pkill -f "main:app" || true
    nohup venv/bin/gunicorn -w 4 -b 0.0.0.0:5000 --timeout 180 main:app > gunicorn.log 2>&1 &
    sleep 3
    echo "✅ Приложение запущено в фоне!"
fi

echo "🔍 Проверка локального ответа сервера:"
curl -I http://127.0.0.1:5000/ || true
EOF

echo ""
echo "🎉 Развертывание завершено успешно!"
echo "Сервис доступен по адресу: http://${TARGET#*@}:5000"
