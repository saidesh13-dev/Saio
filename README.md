# Saio: PDF to Excel Converter

A web app that converts PDF documents into downloadable Excel workbooks while preserving layout, text, and images as closely as possible.

## Run locally

```powershell
.\.venv\Scripts\Activate.ps1
python app.py
```

Open http://127.0.0.1:5000/.

## Deploy publicly with Render

1. Create a GitHub repository and upload this project.
2. Do not upload `.venv`, `uploads`, or `outputs`.
3. In Render, choose **New > Blueprint** and select the repository.
4. Render will read `render.yaml`, build the Docker image, install Tesseract OCR, and start the web service.
5. Open the HTTPS URL provided by Render.

The HTTPS URL can be used on laptops and phones. On supported browsers, users can choose **Install app** or **Add to Home Screen**.

## Deploy with Docker locally

```powershell
docker build -t pdf-to-excel .
docker run --rm -p 8000:8000 pdf-to-excel
```

Open http://127.0.0.1:8000/.
