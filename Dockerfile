# Use official Python image
FROM python:3.10-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

# Set work directory
WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

# Copy source code
COPY . .

# Expose the port Dash will run on
EXPOSE 8050

# Run the Dash app using gunicorn for production
CMD ["gunicorn", "app:server", "-b", "0.0.0.0:8050"]
