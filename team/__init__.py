# external
from flask import Flask, render_template, abort, redirect, url_for
from flask_bootstrap import Bootstrap
from flask_admin.contrib.sqla import ModelView
from flask_admin.contrib.sqla.filters import FilterEqual, FilterLike
import pandas as pd
import datetime as dt
import pathlib
import os

# internal
# from config import app_config
db_uri = "postgresql://rhamana:T13Rb5iH3zXLOuT9gpyabQJoi3IMC0la@dpg-cjhlqdb37aks73d7ago0-a/db_team_f5x6"


from .extensions import db, login_manager, migrate, admin
from team.models import *

config_name="development"

app = Flask(__name__)
app.config['SECRET_KEY'] = "my-secret-key-#$"
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URL')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['DEBUG'] = True
app.config['SQLALCHEMY_ECHO'] = True

Bootstrap(app)
db.init_app(app)
login_manager.init_app(app)
login_manager.login_message = "You must be logged into access this page!"
login_manager.login_view = "auth.login"
# migrate = Migrate(app, db)
migrate.init_app(app, db)
admin.init_app(app)

class UserView(ModelView):
    # column_exclude_list = ['password']
    column_display_pk = False
    can_create = True
    can_edit = True
    can_delete = True
    can_export = True
    create_modal = True

# Country View
class CountryView(ModelView):
    # column_exclude_list = ['password']
    column_display_pk = False
    can_create = True
    can_edit = True
    can_delete = True
    can_export = True
    create_modal = True

    page_size = 25

    column_list = (
        'country',
    )

    column_editable_list = ['country']
    # column_searchable_list = ['name', 'matricule']
    column_filters = ['country',]
    column_filter_labels = {
        'country': "Pays",
    }

    def scaffold_filters(self, name):
        filters = super().scaffold_filters(name)
        if name in self.column_filter_labels:
            for f in filters:
                f.name = self.column_filter_labels[name]
        return filters

    # override the column labels
    column_labels = {
        'country': "Pays",
    }


# Coordinator View
class CoordinatorView(ModelView):
    # column_exclude_list = ['password']
    column_display_pk = False
    can_create = True
    can_edit = True
    can_delete = True
    can_export = True
    create_modal = True

    page_size = 25

    column_list = (
        'matricule',
        'name',
        'hire_date',
        'departure_date',
        'status',
        'salary',
        'bonus',
        'coordination',
        'coordination.country',
    )

    column_editable_list = ['salary', 'bonus', 'departure_date', 'coordination',]
    # column_searchable_list = ['name', 'matricule']
    column_filters = [
        'name',
        'matricule',
        'coordination',
        'coordination.country',
    ]

    column_filter_labels = {
        'name': 'Nom',
        'matricule': "Matricule",
        'coordination': "Coordination",
        'coordination.country': 'Pays',
    }

    def scaffold_filters(self, name):
        filters = super().scaffold_filters(name)
        if name in self.column_filter_labels:
            for f in filters:
                f.name = self.column_filter_labels[name]
        return filters

    # override the column labels
    column_labels = {
        'matricule': "Matricule",
        'name': "Nom",
        'hire_date': "Embauche",
        'departure_date': "Départ",
        'status': "Statut",
        'salary': "Salaire",
        'bonus': "Bonus",
        'coordination': "Coordination",
        'coordination.country': "Pays",
    }

    def on_model_change(self, form, model, is_created):
        if model.departure_date:
            if model.hire_date <= model.departure_date:
                model.status_id = Status.query.filter_by(employee_status="Inactive").first().id
            else:
                model.status_id = Status.query.filter_by(employee_status="Incoherent").first().id


# Coordination View
class CoordinationView(ModelView):
    # column_exclude_list = ['password']
    column_display_pk = False
    can_create = True
    can_edit = True
    can_delete = True
    can_export = True
    create_modal = True

    page_size = 25

    column_list = (
        'coordination',
        'country',
    )

    column_editable_list = ['coordination',]
    # column_searchable_list = ['name', 'matricule']
    column_filters = [
        'coordination',
        'country',
    ]

    column_filter_labels = {
        'coordination': 'Coordination',
        'country': 'Pays',
    }

    def scaffold_filters(self, name):
        filters = super().scaffold_filters(name)
        if name in self.column_filter_labels:
            for f in filters:
                f.name = self.column_filter_labels[name]
        return filters

    # override the column labels
    column_labels = {
        'coordination': "Coordination",
        'country': "Pays",
    }


# Agency Chief View
class AgencyChiefView(ModelView):
    # column_exclude_list = ['password']
    column_display_pk = False
    can_create = True
    can_edit = True
    can_delete = True
    can_export = True
    create_modal = True

    page_size = 25

    column_list = (
        'matricule',
        'name',
        'hire_date',
        'departure_date',
        'status',
        'salary',
        'bonus',
        'agency',
        'agency.coordination',
        'agency.coordination.country',
    )

    column_editable_list = ['salary', 'bonus', 'departure_date', 'agency']
    # column_searchable_list = ['name', 'matricule']
    column_filters = [
        'name',
        'matricule',
        'agency',
        'agency.coordination',
        'agency.coordination.country',
    ]

    column_filter_labels = {
        'name': "Nom",
        'matricule': 'Matricule',
        'agency': 'Agence',
        'agency.coordination': 'Coordination',
        'agency.coordination.country': 'Pays',
    }

    def scaffold_filters(self, name):
        filters = super().scaffold_filters(name)
        if name in self.column_filter_labels:
            for f in filters:
                f.name = self.column_filter_labels[name]
        return filters

    # override the column labels
    column_labels = {
        'matricule': "Matricule",
        'name': "Nom",
        'hire_date': "Embauche",
        'departure_date': "Départ",
        'status': "Statut",
        'salary': "Salaire",
        'bonus': "Bonus",
        'agency': "Agence",
        'agency.coordination': "Coordination",
        'agency.coordination.country': "Pays",
    }

    def on_model_change(self, form, model, is_created):
        if model.departure_date:
            if model.hire_date <= model.departure_date:
                model.status_id = Status.query.filter_by(employee_status="Inactive").first().id
            else:
                model.status_id = Status.query.filter_by(employee_status="Incoherent").first().id


# Agency View
class AgencyView(ModelView):
    # column_exclude_list = ['password']
    column_display_pk = False
    can_create = True
    can_edit = True
    can_delete = True
    can_export = True
    create_modal = True

    page_size = 25

    column_list = (
        'agency',
        'coordination',
        'coordination.country',
    )

    column_editable_list = ['agency',]
    # column_searchable_list = ['name', 'matricule']
    column_filters = [
        'agency',
        'coordination',
        'coordination.country',
    ]

    column_filter_labels = {
        'agency': 'Agence',
        'coordination': 'Coordination',
        'coordination.country': 'Pays',
    }

    def scaffold_filters(self, name):
        filters = super().scaffold_filters(name)
        if name in self.column_filter_labels:
            for f in filters:
                f.name = self.column_filter_labels[name]
        return filters

    # override the column labels
    column_labels = {
        'agency': 'Agence',
        'coordination': "Coordination",
        'coordination.country': "Pays",
    }


# Agent View
class AgentView(ModelView):
    # column_exclude_list = ['password']
    column_display_pk = False
    can_create = True
    can_edit = True
    can_delete = True
    can_export = True
    create_modal = True

    page_size = 25

    column_list = (
        'matricule',
        'name',
        'hire_date',
        'departure_date',
        'status',
        'salary',
        'bonus',
        'pos',
        'pos.agency',
        'pos.agency.coordination',
        'pos.agency.coordination.country',
    )

    column_editable_list = ['salary', 'bonus', 'departure_date', 'pos']
    # column_searchable_list = ['name', 'matricule']
    column_filters = [
        'name',
        'matricule',
        'pos',
        'pos.agency',
        'pos.agency.coordination',
        'pos.agency.coordination.country',
    ]

    column_filter_labels = {
        'name': 'Nom',
        'matricule': 'Matricule',
        'pos': 'PDV',
        'pos.agency': 'Agence',
        'pos.agency.coordination': 'Coordination',
        'pos.agency.coordination.country': 'Pays',
    }

    def scaffold_filters(self, name):
        filters = super().scaffold_filters(name)
        if name in self.column_filter_labels:
            for f in filters:
                f.name = self.column_filter_labels[name]
        return filters

    # override the column labels
    column_labels = {
        'matricule': "Matricule",
        'name': "Nom",
        'hire_date': "Embauche",
        'departure_date': "Départ",
        'status': "Statut",
        'salary': "Salaire",
        'bonus': "Bonus",
        'pos': "PDV",
        'pos.agency': "Agence",
        'pos.agency.coordination': "Coordination",
        'pos.agency.coordination.country': "Pays"
    }

    def on_model_change(self, form, model, is_created):
        if model.departure_date:
            if model.hire_date <= model.departure_date:
                model.status_id = Status.query.filter_by(employee_status="Inactive").first().id
            else:
                model.status_id = Status.query.filter_by(employee_status="Incoherent").first().id


# POS View
class POSView(ModelView):
    # column_exclude_list = ['password']
    column_display_pk = False
    can_create = True
    can_edit = True
    can_delete = True
    can_export = True
    create_modal = True

    page_size = 25

    column_list = (
        'point_of_sale',
        'agency',
        'agency.coordination',
        'agency.coordination.country',
    )

    column_editable_list = ['point_of_sale',]
    # column_searchable_list = ['name', 'matricule']
    column_filters = [
        'point_of_sale',
        'agency',
        'agency.coordination',
        'agency.coordination.country',
    ]

    column_filter_labels = {
        'point_of_sale': 'PDV',
        'agency': 'Agence',
        'agency.coordination': 'Coordination',
        'agency.coordination.country': 'Pays',
    }

    def scaffold_filters(self, name):
        filters = super().scaffold_filters(name)
        if name in self.column_filter_labels:
            for f in filters:
                f.name = self.column_filter_labels[name]
        return filters

    # override the column labels
    column_labels = {
        'point_of_sale': 'PDV',
        'agency': 'Agence',
        'agency.coordination': "Coordination",
        'agency.coordination.country': "Pays",
    }


# admin.add_view(ModelView(User, db.session))
# admin.add_view(ModelView(Employee, db.session, name="Base"))

admin.add_view(AgentView(Agent, db.session, name="Agents guichet"))
admin.add_view(AgencyChiefView(AgencyChief, db.session, name="Chefs d'agence"))
admin.add_view(CoordinatorView(Coordinator, db.session, name="Coordonateurs"))

admin.add_view(CountryView(Country, db.session, name="Pays"))
admin.add_view(CoordinationView(Coordination, db.session, name="Coordonations"))
admin.add_view(AgencyView(Agency, db.session, name="Agences"))
admin.add_view(POSView(PointOfSale, db.session, name="PDVs"))

# admin.add_view(ModelView(Status, db.session))
# admin.add_view(ModelView(Level, db.session))



@app.route('/', methods=['GET', 'POST'])
def index():
    return redirect("/admin")



@app.route('/load_arch', methods=['GET', 'POST'])
def load_arch():
    fake_db_path = pathlib.Path(__file__).parent.joinpath('fake_db.xlsx').resolve()
    archi_df = pd.read_excel(io=fake_db_path, sheet_name='PDVs')

    for idx, row in archi_df.iterrows():

        # Country
        country_ = Country.query.filter_by(country=row['PAYS']).first()
        if not country_:
            country_ = Country(country=row['PAYS'])
            db.session.add(country_)

        # Coord
        coord_ = Coordination.query.filter_by(coordination=row['COORD']).first()
        if not coord_:
            coord_ = Coordination(coordination=row['COORD'], country_id=country_.id)
            db.session.add(coord_)

        # Agence
        agence_ = Agency.query.filter_by(agency=row['AGENCE']).first()
        if not agence_:
            agence_ = Agency(agency=row['AGENCE'], coordination_id=coord_.id)
            db.session.add(agence_)

        # PDV
        pos_ = PointOfSale.query.filter_by(point_of_sale=row['PDV']).first()
        if not pos_:
            pos_ = PointOfSale(point_of_sale=row['PDV'], agency_id=agence_.id)
            db.session.add(pos_)

    db.session.commit()

    return 'OK'


@app.route('/load_employee_data', methods=['GET', 'POST'])
def load_employee_data():
    fake_db_path = pathlib.Path(__file__).parent.joinpath('fake_db.xlsx').resolve()

    # Status
    status_df = pd.read_excel(io=fake_db_path, sheet_name='Status')
    for idx, row in status_df.iterrows():
        status_ = Status.query.filter_by(employee_status=row['Status']).first()
        if not status_:
            status_ = Status(employee_status=row['Status'])
            db.session.add(status_)

    db.session.commit()

    # Levels
    levels_df = pd.read_excel(io=fake_db_path, sheet_name='Levels')
    for idx, row in levels_df.iterrows():
        level_ = Level.query.filter_by(employee_level=row['Level']).first()
        if not level_:
            level_ = Level(employee_level=row['Level'])
            db.session.add(level_)

    db.session.commit()

    return 'OK'



@app.route('/load_employees', methods=['GET', 'POST'])
def load_employees():
    fake_db_path = pathlib.Path(__file__).parent.joinpath('fake_db.xlsx').resolve()

    # Coord & Coord Adjoint
    df_01 = pd.read_excel(io=fake_db_path, sheet_name='Coordonateurs', parse_dates=['DATE_EMBAUCHE'])
    df_02 = pd.read_excel(io=fake_db_path, sheet_name='Coordonateurs Adjoints', parse_dates=['DATE_EMBAUCHE'])
    df_coord = pd.concat([df_01, df_02])
    for idx, row in df_coord.iterrows():
        coord_ = Coordinator.query.filter_by(matricule=row['MAT']).first()
        if not coord_:
            hire_date_ = dt.date(row['DATE_EMBAUCHE'].year, row['DATE_EMBAUCHE'].month, row['DATE_EMBAUCHE'].day)
            coord_ = Coordinator(
                matricule=row['MAT'],
                name=row['EMPLOYEE'],
                hire_date=hire_date_,
                salary=row['SALAIRE'],
                bonus=row['BONUS'],
                level_id=Level.query.filter_by(employee_level=row['CATEGORIE']).first().id,
                status_id=Status.query.filter_by(employee_status="Active").first().id,
                coordination_id=Coordination.query.filter_by(coordination=row['COORD']).first().id
            )

            db.session.add(coord_)

    db.session.commit()


    # Agence
    df_agence = pd.read_excel(io=fake_db_path, sheet_name="Chefs d'agence", parse_dates=['DATE_EMBAUCHE'])
    for idx, row in df_agence.iterrows():
        chief_ = AgencyChief.query.filter_by(matricule=row['MAT']).first()
        if not chief_:
            hire_date_ = dt.date(row['DATE_EMBAUCHE'].year, row['DATE_EMBAUCHE'].month, row['DATE_EMBAUCHE'].day)
            chief_ = AgencyChief(
                matricule=row['MAT'],
                name=row['EMPLOYEE'],
                hire_date=hire_date_,
                salary=row['SALAIRE'],
                bonus=row['BONUS'],
                level_id=Level.query.filter_by(employee_level=row['CATEGORIE']).first().id,
                status_id=Status.query.filter_by(employee_status="Active").first().id,
                agency_id=Agency.query.filter_by(agency=row['AGENCE']).first().id
            )

            db.session.add(chief_)

    db.session.commit()


    # Agent
    df_agents = pd.read_excel(io=fake_db_path, sheet_name="Agents", parse_dates=['DATE_EMBAUCHE'])
    for idx, row in df_agents.iterrows():
        agent_ = Agent.query.filter_by(matricule=row['MAT']).first()
        if not agent_:
            hire_date_ = dt.date(row['DATE_EMBAUCHE'].year, row['DATE_EMBAUCHE'].month, row['DATE_EMBAUCHE'].day)
            agent_ = Agent(
                matricule=row['MAT'],
                name=row['EMPLOYEE'],
                hire_date=hire_date_,
                salary=row['SALAIRE'],
                bonus=row['BONUS'],
                level_id=Level.query.filter_by(employee_level=row['CATEGORIE']).first().id,
                status_id=Status.query.filter_by(employee_status="Active").first().id,
                pos_id=PointOfSale.query.filter_by(point_of_sale=row['PDV']).first().id
            )

            db.session.add(agent_)

    db.session.commit()

    return 'OK'

def rankCat(cat):
    if cat == "Coordonateur":
        return 1
    elif cat == "Coordonateur adjoint":
        return 3
    elif cat == "Agent":
        return 4
    else:
        return 5

@app.route('/export', methods=['GET', 'POST'])
def export():
    # Export Coordonators
    q_coord = Coordinator.query.all()
    if q_coord:
        data_coord = [{
            'Matricule': i.matricule,
            'Nom': i.name,
            'Embauche': i.hire_date,
            'Départ': i.departure_date,
            'Salaire': i.salary,
            'Bonus': i.bonus,
            'Statut': i.status,
            'Catégorie': i.level,
            'Coordination': i.coordination,
            'Pays': i.coordination.country
        } for i in q_coord]

        df_coord = pd.DataFrame(data_coord)
        df_coord['Rang'] = 1
    else:
        df_coord = pd.DataFrame()

    # Export Agency Chiefs
    q_agency = AgencyChief.query.all()
    if q_agency:
        data_agency = [{
            'Matricule': i.matricule,
            'Nom': i.name,
            'Embauche': i.hire_date,
            'Départ': i.departure_date,
            'Salaire': i.salary,
            'Bonus': i.bonus,
            'Statut': i.status,
            'Catégorie': i.level,
            'Agence': i.agency,
            'Coordination': i.agency.coordination,
            'Pays': i.agency.coordination.country
        } for i in q_agency]

        df_agency = pd.DataFrame(data_agency)
        df_agency['Rang'] = 2
    else:
        df_agency = pd.DataFrame()


    # Export Agents
    q_agent = Agent.query.all()
    if q_agent:
        data_agent = [{
            'Matricule': i.matricule,
            'Nom': i.name,
            'Embauche': i.hire_date,
            'Départ': i.departure_date,
            'Salaire': i.salary,
            'Bonus': i.bonus,
            'Statut': i.status,
            'Catégorie': i.level,
            'PDV': i.pos,
            'Agence': i.pos.agency,
            'Coordination': i.pos.agency.coordination,
            'Pays': i.pos.agency.coordination.country
        } for i in q_agent]

        df_agent = pd.DataFrame(data_agent)
        df_agent['Rang'] = 3
    else:
        df_agent = pd.DataFrame()

    out_ = df_coord
    out_ = pd.concat([out_, df_agency])
    out_ = pd.concat([out_, df_agent])

    out_['Pays'] = out_['Pays'].astype(str)
    out_['Coordination'] = out_['Coordination'].astype(str)
    out_['Agence'] = out_['Agence'].astype(str)
    out_ = out_.sort_values(['Pays', 'Coordination', 'Rang'])

    cols_ = [
        'Matricule',
        'Nom',
        'Embauche',
        'Départ',
        'Salaire',
        'Bonus',
        'Statut',
        'Catégorie',
        'Coordination',
        'Pays',
        'Agence',
        'PDV'
    ]

    out_ = out_[cols_].copy()

    out_fake_db_path = pathlib.Path(__file__).parent.joinpath('out_fake_db.xlsx').resolve()
    out_.to_excel(
        excel_writer=out_fake_db_path,
        sheet_name="Out",
        index=False
    )

    return 'OK'
