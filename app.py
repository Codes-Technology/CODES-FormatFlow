"""
Main application file with JWT authentication and Session Reset
"""
from datetime import timedelta
from flask import Flask, jsonify, render_template, redirect, request, url_for, make_response
from flask_jwt_extended import JWTManager, verify_jwt_in_request, get_jwt_identity, get_jwt, unset_jwt_cookies
from flask_cors import CORS
from config import SQLALCHEMY_DATABASE_URI, SQLALCHEMY_TRACK_MODIFICATIONS, APP_INSTANCE_ID 
from utils.db_manager import db, init_db


   
app = Flask(__name__)

CORS(app, supports_credentials=True)

app.config['SQLALCHEMY_DATABASE_URI'] = SQLALCHEMY_DATABASE_URI
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = SQLALCHEMY_TRACK_MODIFICATIONS

# Every restart = new key = old cookies invalid = login required
app.config['SECRET_KEY'] = APP_INSTANCE_ID
app.config['JWT_SECRET_KEY'] = APP_INSTANCE_ID

app.config['JWT_TOKEN_LOCATION'] = ['cookies']
app.config['JWT_COOKIE_HTTPONLY'] = True
app.config['JWT_COOKIE_SECURE'] = False
app.config['JWT_COOKIE_CSRF_PROTECT'] = False
app.config['JWT_ACCESS_TOKEN_EXPIRES'] = timedelta(hours=24)  # ← Bug 2 fix: was int not timedelta

jwt = JWTManager(app)
init_db(app)

from route.auth import auth_bp
from route.process import process_bp
app.register_blueprint(auth_bp, url_prefix='/api/auth')
app.register_blueprint(process_bp, url_prefix='/api/process')


# ── FRONTEND ROUTES ────────────────────────────────────────

@app.route('/login')
def login_page():
    return render_template('login.html')

@app.route('/')
def index():
    try:
        verify_jwt_in_request(optional=True)
        user_id = get_jwt_identity()

        if not user_id:
            return redirect(url_for('login_page'))

        claims = get_jwt()
        token_instance = claims.get('instance_id')
        
        if token_instance != APP_INSTANCE_ID:
            print("❌ Instance mismatch — forcing re-login")
            response = make_response(redirect(url_for('login_page')))
            unset_jwt_cookies(response)
            return response

        response = make_response(render_template('upload.html'))
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '0'
        return response

    except Exception as e:
        print(f"❌ Authentication error: {e}")
        return redirect(url_for('login_page'))


# ── API ROUTES ─────────────────────────────────────────────

@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({'success': True, 'message': 'Prisma API is running', 'version': '1.0.0'}), 200


# ── JWT ERROR HANDLERS ─────────────────────────────────────
# Bug 3 fix: handlers now redirect browser requests, return JSON for API calls

@jwt.expired_token_loader
def expired_token_callback(jwt_header, jwt_payload):
    if request.path.startswith('/api/'):
        return jsonify({'success': False, 'error': 'Token expired'}), 401
    response = make_response(redirect(url_for('login_page')))
    unset_jwt_cookies(response)
    return response

@jwt.invalid_token_loader
def invalid_token_callback(error):
    if request.path.startswith('/api/'):
        return jsonify({'success': False, 'error': 'Invalid token'}), 401
    response = make_response(redirect(url_for('login_page')))
    unset_jwt_cookies(response)
    return response

@jwt.unauthorized_loader
def missing_token_callback(error):
    if request.path.startswith('/api/'):
        return jsonify({'success': False, 'error': 'No token provided'}), 401
    return redirect(url_for('login_page'))


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)