import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

DATA_DIR = os.path.join(BASE_DIR, 'data')
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
PHOTOS_DIR = os.path.join(UPLOAD_DIR, 'photos')

STATIONS_FILE = os.path.join(DATA_DIR, 'stations.json')
HIERARCHY_FILE = os.path.join(BASE_DIR, 'hierarchy.json')

MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB
ALLOWED_PHOTO_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'pdf'}

# 确保目录存在
for d in [DATA_DIR, UPLOAD_DIR, PHOTOS_DIR]:
    os.makedirs(d, exist_ok=True)
