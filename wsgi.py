"""
WSGI入口文件，用于生产环境部署（Gunicorn）
"""
from procurement_agent import app

if __name__ == "__main__":
    app.run()
