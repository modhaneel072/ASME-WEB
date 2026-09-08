release: python manage.py upgrade
web: gunicorn --bind :8000 --workers 2 --timeout 120 app:app
