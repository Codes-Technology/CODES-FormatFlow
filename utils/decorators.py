"""
Custom decorators for route protection
"""
from functools import wraps
from flask import jsonify
from flask_jwt_extended import jwt_required, get_jwt_identity, get_jwt
from utils.db_manager import User

def require_auth(fn):
    @wraps(fn)
    @jwt_required()
    def wrapper(*args, **kwargs):
        try:
            user_id = get_jwt_identity()
            claims = get_jwt()

            user = User.query.get(int(user_id))

            if not user:
                return jsonify({
                    'success': False,
                    'error': 'User not found'
                }), 404
            
            if not user.IsActive:
                return jsonify({
                    'success': False,
                    'error': 'Account disabled'
                }), 403
            
            # IMPORTANT: Check token instance (app restart forced invalidation)
            from config import APP_INSTANCE_ID
            if claims.get('instance_id') != APP_INSTANCE_ID:
                return jsonify({
                    'success': False,
                    'error': 'Session expired due to server restart. Please login again.'
                }), 401
            
            # IMPORTANT: Check token version (user-initiated logout security)
            if claims.get('token_version') != user.TokenVersion:
                return jsonify({
                    'success': False,
                    'error': 'Session expired. Please login again.'
                }), 401
            
            return fn(current_user = user, *args, **kwargs)
        
        except Exception as e:
            return jsonify({
                'success': False,
                'error': f"Authentication failed: {str(e)}"
            }), 401
        
    return wrapper