from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_migrate import Migrate
import os

db = SQLAlchemy()
migrate = Migrate()

def create_app(config_name='default'):
    """Application factory pattern for Flask app creation."""
    
    app = Flask(__name__)
    
    # Load configuration
    app.config.from_object('config.Config')
    
    # Initialize extensions
    db.init_app(app)
    migrate.init_app(app, db)
    
    # Register blueprints
    from app.routes.main import bp as main_bp
    app.register_blueprint(main_bp)
    
    from app.routes.upload import bp as upload_bp
    app.register_blueprint(upload_bp, url_prefix='/upload')
    
    from app.routes.matching import bp as matching_bp
    app.register_blueprint(matching_bp, url_prefix='/matching')
    
    from app.routes.email import bp as email_bp
    app.register_blueprint(email_bp, url_prefix='/email')
    
    from app.routes.api import bp as api_bp
    app.register_blueprint(api_bp, url_prefix='/api')
    
    # Create tables
    with app.app_context():
        db.create_all()
    
    return app