# python -m pip install --upgrade pip
# external
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

# internal
from .extensions import db, login_manager


class User(UserMixin, db.Model):
    """ Create an Employee table"""
    __tablename__ = 'users'

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(60), index=True, unique=True)
    first_name = db.Column(db.String(60), index=True)
    last_name = db.Column(db.String(60), index=True)
    password_hash = db.Column(db.String(128))
    is_admin = db.Column(db.Boolean, default=False)

    @property
    def password(self):
        """Prevent password from being accessed"""
        raise AttributeError('Password is not a readable attribute!')

    @password.setter
    def password(self, password):
        """Set password to a hashed password"""
        self.password_hash = generate_password_hash(password)

    def verify_password(self, password):
        """Check if hashed password matches actual password"""
        return check_password_hash(self.password_hash, password)

    def __repr__(self):
        return '<User: {}>'.format(self.username)


# Set up user_loader
@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


class Status(db.Model):
    __tablename__ = 'status'

    id = db.Column(db.Integer, primary_key=True)
    employee_status = db.Column(db.String, index=True, nullable=False, unique=True)
    employees = db.relationship('Employee', backref='status', lazy='dynamic', enable_typechecks=False)

    def __repr__(self):
        return self.employee_status


class Level(db.Model):
    __tablename__ = 'levels'

    id = db.Column(db.Integer, primary_key=True)
    employee_level = db.Column(db.String, index=True, nullable=False, unique=True)
    employees = db.relationship('Employee', backref='level', lazy='dynamic', enable_typechecks=False)

    def __repr__(self):
        return self.employee_level


class Employee(db.Model):
    __tablename__ = 'employees'

    id = db.Column(db.Integer, primary_key=True)
    matricule = db.Column(db.String, index=True, nullable=False, unique=True)
    name = db.Column(db.String, index=False, nullable=False, unique=False)
    hire_date = db.Column(db.Date, index=False, nullable=True, unique=False)
    departure_date = db.Column(db.Date, index=False, nullable=True, unique=False)
    salary = db.Column(db.Float, index=False, nullable=True, unique=False)
    bonus = db.Column(db.Float, index=False, nullable=True, unique=False)

    level_id = db.Column(db.Integer, db.ForeignKey('levels.id'))
    status_id = db.Column(db.Integer, db.ForeignKey('status.id'))

    def __repr__(self):
        return f"<{self.matricule}, {self.name}>"


class Agent(Employee):
    __tablename__ = 'agents'

    id = db.Column(db.Integer, db.ForeignKey('employees.id'), primary_key=True)
    pos_id = db.Column(db.Integer, db.ForeignKey('pos.id'))


class AgencyChief(Employee):
    __tablename__ = 'agency_chiefs'

    id = db.Column(db.Integer, db.ForeignKey('employees.id'), primary_key=True)
    agency_id = db.Column(db.Integer, db.ForeignKey('agencies.id'))


class Coordinator(Employee):
    __tablename__ = 'coordinators'

    id = db.Column(db.Integer, db.ForeignKey('employees.id'), primary_key=True)
    coordination_id = db.Column(db.Integer, db.ForeignKey('coordinations.id'))


class PointOfSale(db.Model):
    __tablename__ = 'pos'

    id = db.Column(db.Integer, primary_key=True)
    point_of_sale = db.Column(db.String, index=False, nullable=False, unique=True)
    agency_id = db.Column(db.Integer, db.ForeignKey('agencies.id'))
    employees = db.relationship('Agent', backref='pos', lazy='dynamic', enable_typechecks=False)

    def __repr__(self):
        return self.point_of_sale


class Agency(db.Model):
    __tablename__ = 'agencies'

    id = db.Column(db.Integer, primary_key=True)
    agency = db.Column(db.String, index=False, nullable=False, unique=True)
    coordination_id = db.Column(db.Integer, db.ForeignKey('coordinations.id'))
    employees = db.relationship('AgencyChief', backref='agency', lazy='dynamic', enable_typechecks=False)
    pos = db.relationship('PointOfSale', backref='agency', lazy='dynamic', enable_typechecks=False)

    def __repr__(self):
        return self.agency


class Coordination(db.Model):
    __tablename__ = 'coordinations'

    id = db.Column(db.Integer, primary_key=True)
    coordination = db.Column(db.String, index=False, nullable=False, unique=True)
    country_id = db.Column(db.Integer, db.ForeignKey('countries.id'))
    employees = db.relationship('Coordinator', backref='coordination', lazy='dynamic', enable_typechecks=False)
    agencies = db.relationship('Agency', backref='coordination', lazy='dynamic', enable_typechecks=False)

    def __repr__(self):
        return self.coordination


class Country(db.Model):
    __tablename__ = 'countries'

    id = db.Column(db.Integer, primary_key=True)
    country = db.Column(db.String, index=False, nullable=False, unique=True)
    coordinations = db.relationship('Coordination', backref='country', lazy='dynamic', enable_typechecks=False)

    def __repr__(self):
        return self.country
