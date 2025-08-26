# gunicorn.conf.py

# Workers = (2 x $num_cores) + 1 is a common formula
import multiprocessing

workers = multiprocessing.cpu_count() * 2 + 1
threads = 2
timeout = 120
graceful_timeout = 30
loglevel = "info"

accesslog = "logs/gunicorn_access.log"
errorlog = "logs/gunicorn_error.log"

