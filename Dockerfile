FROM python:3.10-slim

# Work directory inside the container
WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the code
COPY . .

# Render will use this port
ENV PORT=10000
EXPOSE 10000

# Start the bot
CMD ["python", "bot.py"]
