# Student Progress Maker

Upload a class record, enter the program/year/section, course code, academic year, course description, course type and term, then download one Word document with a progress sheet for every named student.

## Run with Docker

From this folder:

    docker compose up -d --build --wait

Open http://localhost:8000. For online use, put the container behind an authenticated HTTPS endpoint and set its upload limit to 11 MB. The app processes uploads for each request and does not keep copies. The Docker build excludes the sample class record.

## Update a Docker deployment

From the Git checkout on the Docker host:

    bash deploy/update.sh

The updater requires Docker Compose v2 and a clean checkout. It pulls the tracked branch using a fast-forward update, builds the new image before replacing the running container, and waits for the web app to respond. Set STUDPROG_PORT or STUDPROG_BIND_ADDRESS in a local .env file if your deployment uses a different port or bind address.

## Run with Python

    python -m pip install -r requirements.txt
    python -m uvicorn app:app

## Workbook and template

Use an XLSX or XLSM workbook with one worksheet matching the sample class record. Put the headers in row 1. The app reads StudentName, PrelimExam, MidtermExam, FinalExam, MidtermGrade and columns 1–7 for each of MQ, MA, FQ and FA. Blank scores remain blank. Formula results come from the values saved by Excel. Recalculate and save before uploading a workbook with formulas.

The template's first exam row serves as the column labels. The second row receives the scores. Attendance, experiments, other activities and signatures remain available for manual entry.
