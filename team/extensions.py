from flask_sqlalchemy import SQLAlchemy
from flask_migrate import Migrate
from flask_login import LoginManager
from flask_admin import Admin

# db variable initialization
db = SQLAlchemy()
migrate = Migrate()
login_manager = LoginManager()
admin = Admin(template_mode='bootstrap3')
