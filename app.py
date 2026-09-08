"""WSGI entry point.

``gunicorn app:app`` and Elastic Beanstalk's ``WSGIPath: app:app`` both keep
working; everything else lives in the ``asme`` package.
"""

from asme import create_app

app = create_app()

if __name__ == "__main__":
    settings = app.config["SETTINGS"]
    app.run(host="0.0.0.0", port=settings.port, debug=settings.env == "development")
