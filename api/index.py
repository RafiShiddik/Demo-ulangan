import os
import sys

# Add parent directory to sys.path so app.py can be imported cleanly
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app as flask_app

class VercelPathFixer:
    def __init__(self, app):
        self.app = app

    def __call__(self, environ, start_response):
        path_info = environ.get('PATH_INFO', '')
        for prefix in ['/api/index.py', '/api/index', '/api']:
            if path_info.startswith(prefix):
                path_info = path_info[len(prefix):]
                if not path_info.startswith('/'):
                    path_info = '/' + path_info
                environ['PATH_INFO'] = path_info
                break
        return self.app(environ, start_response)

app = VercelPathFixer(flask_app)
