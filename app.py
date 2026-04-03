import os
import time
import shutil
from flask import Flask, render_template

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'exproprieri-tools-secret-2026')
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
for d in ['uploads', 'output']:
    os.makedirs(os.path.join(BASE_DIR, d), exist_ok=True)

# Register blueprints
from modules.br_generator.routes import bp as br_bp
from modules.verificare_hpv.routes import bp as hpv_bp
from modules.organizare_dosare.routes import bp as org_bp
from modules.doc_cadastrale.routes import bp as doccad_bp
app.register_blueprint(br_bp)
app.register_blueprint(hpv_bp)
app.register_blueprint(org_bp)
app.register_blueprint(doccad_bp)


def cleanup_old_sessions():
    cutoff = time.time() - 86400
    for base in ['uploads', 'output']:
        path = os.path.join(BASE_DIR, base)
        if not os.path.exists(path):
            continue
        for d in os.listdir(path):
            dp = os.path.join(path, d)
            if os.path.isdir(dp) and os.path.getmtime(dp) < cutoff:
                shutil.rmtree(dp, ignore_errors=True)


@app.route('/')
def home():
    cleanup_old_sessions()
    return render_template('home.html')


if __name__ == '__main__':
    print('\n' + '=' * 50)
    print('  Instrumente Exproprieri')
    print('=' * 50)
    print(f'  Local:  http://localhost:5050')
    print('=' * 50 + '\n')
    app.run(host='0.0.0.0', port=5050, debug=False)
