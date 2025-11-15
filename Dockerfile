# Use Python 3.10
FROM python:3.10-slim

# Create app directory
WORKDIR /app

# Copy requirements
COPY requirements.txt /app/

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy bot code
COPY . /app/

# Start bot
CMD ["python", "bot.py"]
