FROM python:3.11-slim

# LibreOffice (تحويل إكسيل لـ PDF) + poppler-utils (تحويل PDF لصور) + خط عربي
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    poppler-utils \
    fonts-freefont-ttf \
    libraqm0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

CMD ["sh", "-c", "gunicorn app:app --bind 0.0.0.0:$PORT --timeout 600 --workers 2"]
