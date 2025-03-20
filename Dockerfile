# Use the official lightweight Python image.
FROM python:3.9-slim

# Set the working directory in the container.
WORKDIR /app

# Copy the requirements file and install dependencies.
COPY requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of your app's source code.
COPY . .

# Expose the port the app runs on.
EXPOSE 8080

# Run the application.
CMD ["python", "app.py"]
