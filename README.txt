UNISPHERE EXCEL PROJECT
=======================

This package uses the exact uploaded App.jsx as the frontend.

LOCAL SETUP
-----------

Terminal 1:
cd "C:\Users\aditi\Downloads\UniSphere-Excel-Render"
py -m pip install -r requirements.txt
py -m uvicorn server:app --reload

Terminal 2:
cd "C:\Users\aditi\Downloads\UniSphere-Excel-Render"
npm install
npm run dev

Open:
http://localhost:5173

Backend health:
http://127.0.0.1:8000/api/health

IMPORTANT
---------
Do not use the old TechZen backend/frontend folders.
This is a standalone Vite + React frontend with the FastAPI Excel backend.

Demo:
nikki@unisphere.edu / 1234
priya@unisphere.edu / 1234
admin@unisphere.edu / 1234
