FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 5000
# --workers MUST stay at 1: job state (`jobs = {}` in jobstore.py) lives in
# one process's memory with no shared store (no Redis/DB). Gunicorn workers
# are separate OS processes, each with its own independent copy of that dict
# — a second worker would round-robin /process and /status across two job
# stores that don't know about each other, causing "Job not found" for a job
# that landed on the other worker. Threads DO share memory within one
# process, so concurrency comes from --threads instead.
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--timeout", "300", "--workers", "1", "--threads", "8", "app:app"]
