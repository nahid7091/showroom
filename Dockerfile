FROM python:3.11-slim

WORKDIR /app

# نصب ابزارهای پایه
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# اول فقط requirements کپی و نصب می‌شود تا کش شود!
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# حالا بقیه سورس‌کد کپی می‌شود
COPY . .

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "80"]
