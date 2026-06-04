from __future__ import annotations

import csv
import io
import os
import secrets
import smtplib
import string
import uuid
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

from flask import Flask, abort, flash, g, redirect, render_template, request, send_file, send_from_directory, session, url_for
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Image as RLImage, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from sqlalchemy import inspect, text, or_
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent


def load_local_env() -> None:
    env_path = BASE_DIR / '.env'
    if not env_path.exists():
        return
    for encoding in ('utf-8-sig', 'utf-8', 'cp1252', 'utf-16'):
        try:
            raw_lines = env_path.read_text(encoding=encoding).splitlines()
            break
        except UnicodeDecodeError:
            continue
    else:
        return
    for raw_line in raw_lines:
        line = raw_line.strip()
        if not line or line.startswith('#') or '=' not in line:
            continue
        key, value = line.split('=', 1)
        key = key.strip().lstrip('\ufeff')
        value = value.strip().strip('\"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


load_local_env()
UPLOAD_DIR = BASE_DIR / 'static' / 'uploads'
ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "gif", "webp", "pdf", "txt", "log", "mp4"}

db = SQLAlchemy()


ROLE_LABELS = {
    'admin': 'Administrador',
    'manager': 'Gestor de proyectos',
    'tester': 'Tester / Analista QA',
    'viewer': 'Consulta'
}
ROLE_DEFAULTS = {
    'admin': dict(can_create_projects=True, can_join_projects=True, can_manage_records=True, can_manage_users=True, can_view_audit=True),
    'manager': dict(can_create_projects=True, can_join_projects=True, can_manage_records=True, can_manage_users=False, can_view_audit=True),
    'tester': dict(can_create_projects=False, can_join_projects=True, can_manage_records=True, can_manage_users=False, can_view_audit=False),
    'viewer': dict(can_create_projects=False, can_join_projects=False, can_manage_records=False, can_manage_users=False, can_view_audit=False),
}


class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), nullable=False, unique=True)
    full_name = db.Column(db.String(160), nullable=False)
    email = db.Column(db.String(180), nullable=True)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(40), nullable=False, default='tester')
    is_active = db.Column(db.Boolean, nullable=False, default=True)
    must_change_password = db.Column(db.Boolean, nullable=False, default=False)
    can_create_projects = db.Column(db.Boolean, nullable=False, default=False)
    can_join_projects = db.Column(db.Boolean, nullable=False, default=True)
    can_manage_records = db.Column(db.Boolean, nullable=False, default=True)
    can_manage_users = db.Column(db.Boolean, nullable=False, default=False)
    can_view_audit = db.Column(db.Boolean, nullable=False, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    memberships = db.relationship('ProjectMember', backref='user', lazy=True, cascade='all, delete-orphan')

    def set_password(self, password: str) -> None:
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)

    @property
    def is_admin(self) -> bool:
        return self.role == 'admin'

    @property
    def display_name(self) -> str:
        return self.full_name or self.username


class ProjectMember(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    project_role = db.Column(db.String(40), nullable=False, default='Analista QA')
    can_register_records = db.Column(db.Boolean, nullable=False, default=True)
    joined_at = db.Column(db.DateTime, default=datetime.utcnow)
    __table_args__ = (db.UniqueConstraint('project_id', 'user_id', name='uq_project_user'),)


class AuditLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    action = db.Column(db.String(80), nullable=False)
    entity_type = db.Column(db.String(80), nullable=False)
    entity_id = db.Column(db.Integer, nullable=True)
    detail = db.Column(db.Text, nullable=True)
    ip_address = db.Column(db.String(80), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    user = db.relationship('User')


class Project(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False, unique=True)
    description = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    created_by_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    creator = db.relationship('User', foreign_keys=[created_by_id])
    members = db.relationship('ProjectMember', backref='project', lazy=True, cascade='all, delete-orphan')
    test_records = db.relationship('TestRecord', backref='project', lazy=True, cascade='all, delete-orphan')


class TestRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    application_name = db.Column(db.String(120), nullable=False)
    module_name = db.Column(db.String(120), nullable=True)
    version = db.Column(db.String(80), nullable=True)
    environment = db.Column(db.String(120), nullable=True)
    platform = db.Column(db.String(120), nullable=True)
    browser = db.Column(db.String(120), nullable=True)
    test_type = db.Column(db.String(50), nullable=False, default='Funcional')
    status = db.Column(db.String(50), nullable=False, default='Reportada')
    severity = db.Column(db.String(50), nullable=False, default='Media')
    priority = db.Column(db.String(50), nullable=False, default='Media')
    expected_result = db.Column(db.Text, nullable=True)
    actual_result = db.Column(db.Text, nullable=True)
    validation_notes = db.Column(db.Text, nullable=True)
    responsible = db.Column(db.String(120), nullable=True)
    execution_date = db.Column(db.DateTime, nullable=True)
    resolved_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    created_by_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    updated_by_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    creator = db.relationship('User', foreign_keys=[created_by_id])
    updater = db.relationship('User', foreign_keys=[updated_by_id])
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False)
    evidences = db.relationship('Evidence', backref='test_record', lazy=True, cascade='all, delete-orphan')


class Evidence(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    original_name = db.Column(db.String(255), nullable=False)
    file_name = db.Column(db.String(255), nullable=False, unique=True)
    file_type = db.Column(db.String(50), nullable=True)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    uploaded_by_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    uploader = db.relationship('User')
    test_record_id = db.Column(db.Integer, db.ForeignKey('test_record.id'), nullable=False)


STATUS_OPTIONS = ['Reportada', 'Subsanada', 'Descartada']
SEVERITY_OPTIONS = ['Baja', 'Media', 'Alta', 'Crítica']
PRIORITY_OPTIONS = ['Baja', 'Media', 'Alta', 'Urgente']
TEST_TYPE_OPTIONS = ['Funcional', 'Regresión', 'Usabilidad', 'Seguridad', 'Integración', 'Rendimiento']
RESPONSIBLE_OPTIONS = ['QA interno', 'Proveedor', 'Desarrollo interno', 'Mesa de ayuda', 'Usuario solicitante']
PROJECT_ROLE_OPTIONS = ['Administrador del proyecto', 'Líder funcional', 'Analista QA', 'Desarrollador', 'Proveedor']
CLOSED_STATUSES = {'Subsanada', 'Descartada'}


def create_app() -> Flask:
    app = Flask(__name__)
    app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-secret-key')
    app.config['SQLALCHEMY_DATABASE_URI'] = f"sqlite:///{BASE_DIR / 'instance' / 'qa_novedades.db'}"
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
    app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024
    app.config['SESSION_COOKIE_HTTPONLY'] = True
    app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
    db.init_app(app)

    (BASE_DIR / 'instance').mkdir(exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


    @app.before_request
    def load_logged_user():
        g.current_user = None
        user_id = session.get('user_id')
        if user_id:
            g.current_user = User.query.get(user_id)
            if g.current_user and not g.current_user.is_active:
                session.clear()
                g.current_user = None
            if g.current_user and g.current_user.must_change_password:
                allowed_endpoints = {'change_password', 'logout', 'static'}
                if request.endpoint not in allowed_endpoints:
                    return redirect(url_for('change_password'))

    @app.context_processor
    def inject_user():
        return {
            'current_user': g.get('current_user'),
            'can_user_manage_record': can_user_manage_record,
            'can_user_create_record_in_project': can_user_create_record_in_project,
            'can_user_edit_project_details': can_user_edit_project_details,
            'can_user_manage_project_members': can_user_manage_project_members,
            'can_user_remove_project_member': can_user_remove_project_member,
        }

    @app.errorhandler(403)
    def forbidden_error(error):
        flash('No tienes permisos para realizar esta acción. Si necesitas acceso, solicita al administrador que actualice tu rol o tu asociación al proyecto.', 'warning')
        back_url = request.referrer or url_for('index')
        return redirect(back_url)

    @app.errorhandler(404)
    def not_found_error(error):
        flash('El recurso solicitado no existe o ya no está disponible.', 'warning')
        return redirect(url_for('index'))

    @app.route('/login', methods=['GET', 'POST'])
    def login():
        if request.method == 'POST':
            login_email = request.form.get('username', '').strip()
            password = request.form.get('password', '')
            user = User.query.filter(
                or_(
                    db.func.lower(User.username) == login_email.lower(),
                    db.func.lower(User.email) == login_email.lower()
                )
            ).first()
            if user and user.is_active and user.check_password(password):
                session.clear()
                session['user_id'] = user.id
                log_action('login', 'User', user.id, 'Inicio de sesión')
                if user.must_change_password:
                    flash('Debes cambiar tu contraseña para continuar.', 'warning')
                    return redirect(url_for('change_password'))
                flash('Sesión iniciada correctamente.', 'success')
                return redirect(request.args.get('next') or url_for('index'))
            flash('Usuario o contraseña inválidos.', 'danger')
        return render_template('login.html')


    @app.route('/reset-password', methods=['GET', 'POST'])
    def reset_password_request():
        if request.method == 'POST':
            email = request.form.get('email', '').strip().lower()
            user = User.query.filter(or_(db.func.lower(User.username) == email, db.func.lower(User.email) == email)).first()
            if not email:
                flash('Ingresa tu correo electrónico.', 'warning')
                return render_template('reset_password.html')
            if not user or not user.is_active:
                flash('Si el correo existe y está activo, se intentará enviar una nueva clave temporal.', 'info')
                return redirect(url_for('login'))
            temp_password = generate_temporary_password()
            user.set_password(temp_password)
            user.must_change_password = True
            db.session.commit()
            sent, detail = send_welcome_email(user, temp_password, purpose='reset')
            log_action('password_reset_request', 'User', user.id, f'Restablecimiento desde login: {detail}')
            if sent:
                flash('Se envió una nueva clave temporal al correo registrado.', 'success')
                return redirect(url_for('login'))
            flash(f'No se pudo enviar la nueva clave temporal. Detalle SMTP: {detail}', 'danger')
        return render_template('reset_password.html')

    @app.route('/profile', methods=['GET', 'POST'])
    @login_required
    def profile():
        if request.method == 'POST':
            action = request.form.get('action')
            if action == 'send_temp_password':
                if not g.current_user.email:
                    flash('Tu perfil no tiene correo registrado.', 'warning')
                    return redirect(url_for('profile'))
                temp_password = generate_temporary_password()
                g.current_user.set_password(temp_password)
                g.current_user.must_change_password = True
                db.session.commit()
                sent, detail = send_welcome_email(g.current_user, temp_password, purpose='reset')
                log_action('self_temp_password', 'User', g.current_user.id, f'Nueva clave temporal desde perfil: {detail}')
                if sent:
                    session.clear()
                    flash('Se envió una nueva clave temporal a tu correo. Ingresa nuevamente y cámbiala.', 'success')
                    return redirect(url_for('login'))
                flash(f'No se pudo enviar la clave temporal. Detalle SMTP: {detail}', 'danger')
        return render_template('profile.html')

    @app.route('/logout')
    def logout():
        if g.current_user:
            log_action('logout', 'User', g.current_user.id, 'Cierre de sesión')
        session.clear()
        flash('Sesión cerrada.', 'info')
        return redirect(url_for('login'))

    @app.route('/change-password', methods=['GET', 'POST'])
    @login_required
    def change_password():
        if request.method == 'POST':
            current_password = request.form.get('current_password', '')
            new_password = request.form.get('new_password', '')
            confirm_password = request.form.get('confirm_password', '')
            if not g.current_user.check_password(current_password):
                flash('La contraseña actual no es correcta.', 'danger')
                return render_template('change_password.html')
            if len(new_password) < 8:
                flash('La nueva contraseña debe tener mínimo 8 caracteres.', 'danger')
                return render_template('change_password.html')
            if new_password != confirm_password:
                flash('La confirmación no coincide.', 'danger')
                return render_template('change_password.html')
            g.current_user.set_password(new_password)
            g.current_user.must_change_password = False
            db.session.commit()
            log_action('change_password', 'User', g.current_user.id, 'Contraseña actualizada por el usuario')
            flash('Contraseña actualizada correctamente.', 'success')
            return redirect(url_for('index'))
        return render_template('change_password.html')

    @app.context_processor
    def inject_constants():
        return {
            'STATUS_OPTIONS': STATUS_OPTIONS,
            'SEVERITY_OPTIONS': SEVERITY_OPTIONS,
            'PRIORITY_OPTIONS': PRIORITY_OPTIONS,
            'TEST_TYPE_OPTIONS': TEST_TYPE_OPTIONS,
            'RESPONSIBLE_OPTIONS': RESPONSIBLE_OPTIONS,
            'CLOSED_STATUSES': CLOSED_STATUSES,
            'ROLE_LABELS': ROLE_LABELS,
            'PROJECT_ROLE_OPTIONS': PROJECT_ROLE_OPTIONS,
        }

    @app.route('/')
    @login_required
    def index():
        projects = visible_projects_query().order_by(Project.created_at.desc()).all()
        recent_records = visible_records_query().order_by(TestRecord.created_at.desc()).limit(8).all()
        visible_records_base = visible_records_query()
        total_records = visible_records_base.count()
        open_count = visible_records_query().filter(TestRecord.status.notin_(tuple(CLOSED_STATUSES))).count()
        provider_count = visible_records_query().filter(
            TestRecord.status.notin_(tuple(CLOSED_STATUSES)),
            TestRecord.responsible == 'Proveedor'
        ).count()
        closed_count = visible_records_query().filter(TestRecord.status.in_(tuple(CLOSED_STATUSES))).count()
        high_severity_count = visible_records_query().filter(
            TestRecord.status.notin_(tuple(CLOSED_STATUSES)),
            TestRecord.severity.in_(('Alta', 'Crítica'))
        ).count()
        metrics = {
            'projects': visible_projects_query().count(),
            'records': total_records,
            'open': open_count,
            'provider': provider_count,
            'closed': closed_count,
            'high_severity': high_severity_count,
        }
        return render_template('index.html', projects=projects, records=recent_records, metrics=metrics)

    @app.route('/projects', methods=['GET', 'POST'])
    @login_required
    def projects():
        if request.method == 'POST':
            if not (g.current_user.is_admin or g.current_user.can_create_projects):
                abort(403)
            name = request.form.get('name', '').strip()
            description = request.form.get('description', '').strip()
            if not name:
                flash('El nombre del proyecto es obligatorio.', 'danger')
                return redirect(url_for('projects'))
            existing_project = Project.query.filter(db.func.lower(Project.name) == name.lower()).first()
            if existing_project:
                flash('Ya existe un proyecto con ese nombre.', 'warning')
                return redirect(url_for('projects'))
            project = Project(name=name, description=description, created_by_id=g.current_user.id)
            db.session.add(project)
            db.session.flush()
            db.session.add(ProjectMember(project_id=project.id, user_id=g.current_user.id, project_role='Administrador del proyecto', can_register_records=bool(g.current_user.can_manage_records or g.current_user.is_admin)))
            db.session.commit()
            log_action('create', 'Project', project.id, f'Proyecto creado: {project.name}')
            flash('Proyecto creado correctamente.', 'success')
            return redirect(url_for('project_detail', project_id=project.id))

        all_projects = visible_projects_query().order_by(Project.name.asc()).all()
        return render_template('projects.html', projects=all_projects)

    @app.route('/projects/<int:project_id>')
    @login_required
    def project_detail(project_id: int):
        project = Project.query.get_or_404(project_id)
        require_project_access(project)
        status = request.args.get('status', '').strip()
        query = TestRecord.query.filter_by(project_id=project_id).order_by(TestRecord.created_at.desc())
        if status:
            query = query.filter_by(status=status)
        items = query.all()
        all_records = list(project.test_records)
        open_records = [r for r in all_records if r.status not in CLOSED_STATUSES]
        pending_vendor = [
            r for r in open_records
            if (r.responsible or '').strip().lower() in {'proveedor', 'vendor'}
            or 'proveedor' in (r.responsible or '').strip().lower()
        ]
        metrics = {
            'total': len(all_records),
            'open': len(open_records),
            'pending_vendor': len(pending_vendor),
            'closed': len([r for r in all_records if r.status in CLOSED_STATUSES]),
            'subsanadas': len([r for r in all_records if r.status == 'Subsanada']),
            'discarded': len([r for r in all_records if r.status == 'Descartada']),
            'high_severity': len([r for r in open_records if r.severity in ('Alta', 'Crítica')]),
        }
        return render_template('project_detail.html', project=project, records=items, current_status=status, metrics=metrics)

    @app.route('/records')
    @login_required
    def records():
        # El formulario de filtros envía project_id vacío cuando se selecciona "Todos".
        # Parsearlo de forma explícita evita errores 500 ante valores vacíos o no numéricos
        # como /records?project_id=&status=Subsanada.
        raw_project_id = (request.args.get('project_id') or '').strip()
        project_id = int(raw_project_id) if raw_project_id.isdigit() else None

        status = (request.args.get('status') or '').strip()
        if status not in STATUS_OPTIONS:
            status = ''

        if project_id is not None:
            return redirect(url_for('project_detail', project_id=project_id, status=status))

        query = visible_records_query().order_by(TestRecord.created_at.desc())
        if status:
            query = query.filter_by(status=status)

        items = query.all()
        projects = visible_projects_query().order_by(Project.name.asc()).all()
        return render_template('records.html', records=items, projects=projects, current_project_id=project_id, current_status=status)

    @app.route('/projects/<int:project_id>/records/new', methods=['GET', 'POST'])
    @login_required
    def project_record_create(project_id: int):
        project = Project.query.get_or_404(project_id)
        require_record_permission(project)
        return handle_record_form(project=project)

    @app.route('/records/new', methods=['GET', 'POST'])
    @login_required
    def record_create():
        projects = recordable_projects_query().order_by(Project.name.asc()).all()
        if not projects:
            flash('Primero debes crear un proyecto.', 'warning')
            return redirect(url_for('projects'))
        default_project = Project.query.get(request.args.get('project_id', type=int)) if request.args.get('project_id', type=int) else None
        return handle_record_form(project=default_project, projects=projects)

    @app.route('/records/<int:record_id>')
    @login_required
    def record_detail(record_id: int):
        record = TestRecord.query.get_or_404(record_id)
        require_project_access(record.project)
        return render_template('record_detail.html', record=record)

    @app.route('/records/<int:record_id>/edit', methods=['GET', 'POST'])
    @login_required
    def record_edit(record_id: int):
        record = TestRecord.query.get_or_404(record_id)
        require_record_permission(record.project)
        if is_record_closed(record):
            flash('Esta novedad ya está cerrada como subsanada o descartada y no permite nuevas modificaciones.', 'warning')
            return redirect(url_for('record_detail', record_id=record.id))
        projects = recordable_projects_query().order_by(Project.name.asc()).all()

        if request.method == 'POST':
            update_record_from_form(record)
            if not record.module_name or not record.environment or not record.browser or not record.expected_result or not record.actual_result or not record.responsible or not record.execution_date:
                flash('Todos los campos de texto y la fecha de ejecución son obligatorios. Las evidencias son opcionales.', 'danger')
                return render_template('record_form.html', projects=projects, record=record, selected_project=record.project)
            require_record_permission(Project.query.get(record.project_id))
            record.updated_by_id = g.current_user.id
            db.session.commit()
            log_action('update', 'TestRecord', record.id, 'Novedad actualizada')
            save_uploaded_files(request.files.getlist('evidences'), record.id)
            flash('Registro actualizado correctamente.', 'success')
            return redirect(url_for('record_detail', record_id=record.id))

        return render_template('record_form.html', projects=projects, record=record, selected_project=record.project)

    @app.route('/records/<int:record_id>/status', methods=['POST'])
    @login_required
    def record_status_update(record_id: int):
        record = TestRecord.query.get_or_404(record_id)
        require_record_permission(record.project)
        if is_record_closed(record):
            flash('Esta novedad ya está cerrada y no permite cambio de estado.', 'warning')
            return redirect(url_for('record_detail', record_id=record.id))
        new_status = request.form.get('status', '').strip()
        validation_notes = request.form.get('validation_notes', '').strip()
        if new_status not in STATUS_OPTIONS:
            flash('Selecciona un estado válido.', 'danger')
            return redirect(url_for('record_detail', record_id=record.id))
        record.status = new_status
        if validation_notes:
            record.validation_notes = validation_notes
        if new_status in CLOSED_STATUSES and not record.resolved_at:
            record.resolved_at = datetime.utcnow()
        if new_status not in CLOSED_STATUSES:
            record.resolved_at = None
        record.updated_by_id = g.current_user.id
        db.session.commit()
        log_action('status', 'TestRecord', record.id, f'Estado cambiado a {new_status}')
        flash('Estado de la novedad actualizado.', 'success')
        return redirect(url_for('record_detail', record_id=record.id))

    @app.route('/records/<int:record_id>/evidences/add', methods=['POST'])
    @login_required
    def record_evidence_add(record_id: int):
        record = TestRecord.query.get_or_404(record_id)
        require_record_permission(record.project)
        if is_record_closed(record):
            flash('Esta novedad ya está cerrada y no permite agregar evidencias.', 'warning')
            return redirect(url_for('record_detail', record_id=record.id))
        files = request.files.getlist('evidences')
        if not files:
            flash('No se recibieron archivos para cargar.', 'warning')
            return redirect(url_for('record_detail', record_id=record.id))
        save_uploaded_files(files, record.id)
        log_action('upload_evidence', 'TestRecord', record.id, 'Evidencias agregadas')
        flash('Evidencias cargadas correctamente.', 'success')
        return redirect(url_for('record_detail', record_id=record.id))

    @app.route('/projects/<int:project_id>/report.csv')
    @login_required
    def project_report_csv(project_id: int):
        project = Project.query.get_or_404(project_id)
        require_project_access(project)
        mem = build_project_csv(project)
        safe_name = safe_project_name(project)
        return send_file(mem, mimetype='text/csv', as_attachment=True, download_name=f'reporte_novedades_{safe_name}.csv')

    @app.route('/projects/<int:project_id>/report.xlsx')
    @login_required
    def project_report_xlsx(project_id: int):
        project = Project.query.get_or_404(project_id)
        require_project_access(project)
        mem = build_project_xlsx(project)
        safe_name = safe_project_name(project)
        return send_file(
            mem,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'reporte_novedades_{safe_name}.xlsx',
        )

    @app.route('/projects/<int:project_id>/report.pdf')
    @login_required
    def project_report_pdf(project_id: int):
        project = Project.query.get_or_404(project_id)
        require_project_access(project)
        mem = build_project_pdf(project)
        safe_name = safe_project_name(project)
        return send_file(mem, mimetype='application/pdf', as_attachment=True, download_name=f'reporte_ejecutivo_{safe_name}.pdf')

    @app.route('/evidence/<path:filename>')
    @login_required
    def evidence_file(filename: str):
        return send_from_directory(UPLOAD_DIR, filename, as_attachment=False)

    @app.route('/evidence/<int:evidence_id>/delete', methods=['POST'])
    @login_required
    def evidence_delete(evidence_id: int):
        evidence = Evidence.query.get_or_404(evidence_id)
        require_record_permission(evidence.test_record.project)
        if is_record_closed(evidence.test_record):
            flash('Esta novedad ya está cerrada y no permite eliminar evidencias.', 'warning')
            return redirect(url_for('record_detail', record_id=evidence.test_record_id))
        record_id = evidence.test_record_id
        file_path = UPLOAD_DIR / evidence.file_name
        if file_path.exists():
            file_path.unlink()
        db.session.delete(evidence)
        db.session.commit()
        log_action('delete_evidence', 'Evidence', evidence_id, 'Evidencia eliminada')
        flash('Evidencia eliminada.', 'info')
        return redirect(url_for('record_detail', record_id=record_id))



    @app.route('/projects/<int:project_id>/edit', methods=['GET', 'POST'])
    @login_required
    def project_edit(project_id: int):
        project = Project.query.get_or_404(project_id)
        require_project_owner(project)
        if request.method == 'POST':
            old_name = project.name
            project.name = request.form.get('name', '').strip()
            project.description = request.form.get('description', '').strip()
            if not project.name:
                flash('El nombre del proyecto es obligatorio.', 'danger')
                return render_template('project_edit.html', project=project)
            duplicate = Project.query.filter(db.func.lower(Project.name) == project.name.lower(), Project.id != project.id).first()
            if duplicate:
                flash('Ya existe otro proyecto con ese nombre.', 'warning')
                return render_template('project_edit.html', project=project)
            db.session.commit()
            log_action('update', 'Project', project.id, f'Proyecto actualizado: {old_name} -> {project.name}')
            flash('Proyecto actualizado.', 'success')
            return redirect(url_for('project_detail', project_id=project.id))
        return render_template('project_edit.html', project=project)

    @app.route('/projects/<int:project_id>/delete', methods=['POST'])
    @login_required
    def project_delete(project_id: int):
        project = Project.query.get_or_404(project_id)

        if not (g.current_user.is_admin or project.created_by_id == g.current_user.id):
            flash('Solo puedes eliminar proyectos creados por ti. Si el proyecto fue creado por otro usuario, debe eliminarlo un administrador.', 'warning')
            return redirect(url_for('project_detail', project_id=project.id))

        confirm_password = request.form.get('confirm_password', '')
        if not confirm_password or not check_password_hash(g.current_user.password_hash, confirm_password):
            flash('No se eliminó el proyecto. La contraseña ingresada no es válida.', 'danger')
            return redirect(url_for('project_detail', project_id=project.id))

        project_name = project.name
        db.session.delete(project)
        db.session.commit()
        log_action('delete', 'Project', project_id, f'Proyecto eliminado: {project_name}')
        flash(f'Proyecto "{project_name}" eliminado correctamente.', 'success')
        return redirect(url_for('projects'))


    @app.route('/projects/<int:project_id>/join', methods=['POST'])
    @login_required
    def project_join(project_id: int):
        project = Project.query.get_or_404(project_id)
        if not (g.current_user.is_admin or g.current_user.can_join_projects):
            abort(403)
        if not is_project_member(project.id, g.current_user.id):
            db.session.add(ProjectMember(project_id=project.id, user_id=g.current_user.id, project_role='Analista QA', can_register_records=g.current_user.can_manage_records))
            db.session.commit()
            log_action('join', 'Project', project.id, f'Usuario se asoció al proyecto: {project.name}')
            flash('Te asociaste al proyecto correctamente.', 'success')
        return redirect(url_for('project_detail', project_id=project.id))

    @app.route('/projects/<int:project_id>/members', methods=['GET', 'POST'])
    @login_required
    def project_members(project_id: int):
        project = Project.query.get_or_404(project_id)
        require_project_member_admin(project)
        if request.method == 'POST':
            user_id = request.form.get('user_id', type=int)
            can_register = bool(request.form.get('can_register_records'))
            project_role = request.form.get('project_role', 'Analista QA').strip() or 'Analista QA'
            if project_role not in PROJECT_ROLE_OPTIONS:
                project_role = 'Analista QA'
            user = User.query.get_or_404(user_id)
            if not can_user_assign_project_role(project, user, project_role):
                flash('No puedes asignar o modificar ese rol. Solo el creador del proyecto o un administrador global pueden asignar administradores del proyecto o modificar usuarios de mayor perfil.', 'warning')
                return redirect(url_for('project_members', project_id=project.id))
            membership = ProjectMember.query.filter_by(project_id=project.id, user_id=user.id).first()
            if not membership:
                membership = ProjectMember(project_id=project.id, user_id=user.id)
                db.session.add(membership)
            membership.project_role = project_role
            membership.can_register_records = can_register
            db.session.commit()
            log_action('assign_user', 'Project', project.id, f'{user.username} asociado al proyecto')
            flash('Usuario asociado/actualizado en el proyecto.', 'success')
            return redirect(url_for('project_members', project_id=project.id))
        users = User.query.order_by(User.full_name.asc()).all()
        return render_template('project_members.html', project=project, users=users)

    @app.route('/projects/<int:project_id>/members/<int:membership_id>/remove', methods=['POST'])
    @login_required
    def project_member_remove(project_id: int, membership_id: int):
        project = Project.query.get_or_404(project_id)
        require_project_member_admin(project)
        membership = ProjectMember.query.get_or_404(membership_id)
        if membership.project_id != project.id:
            abort(404)
        if not can_user_remove_project_member(project, membership):
            flash('No puedes retirar a este usuario. Un perfil menor no puede eliminar al creador, a un administrador global ni a otro administrador del proyecto.', 'warning')
            return redirect(url_for('project_members', project_id=project.id))
        username = membership.user.username
        db.session.delete(membership)
        db.session.commit()
        log_action('remove_user', 'Project', project.id, f'{username} retirado del proyecto')
        flash('Usuario retirado del proyecto.', 'info')
        return redirect(url_for('project_members', project_id=project.id))

    @app.route('/admin/users', methods=['GET', 'POST'])
    @login_required
    def admin_users():
        require_user_admin()
        if request.method == 'POST':
            email = request.form.get('email', '').strip().lower()
            username = email
            full_name = request.form.get('full_name', '').strip() or email
            role = request.form.get('role', 'tester')
            if not email:
                flash('El correo electrónico es obligatorio. Ese correo será también el usuario de ingreso.', 'danger')
                return redirect(url_for('admin_users'))
            if User.query.filter(or_(db.func.lower(User.username) == email.lower(), db.func.lower(User.email) == email.lower())).first():
                flash('Ya existe un usuario registrado con ese correo.', 'warning')
                return redirect(url_for('admin_users'))
            defaults = ROLE_DEFAULTS.get(role, ROLE_DEFAULTS['tester'])
            temp_password = generate_temporary_password()
            user = User(username=username, full_name=full_name, email=email, role=role, must_change_password=True, **defaults)
            user.set_password(temp_password)
            db.session.add(user)
            db.session.commit()
            sent, detail = send_welcome_email(user, temp_password)
            log_action('create', 'User', user.id, f'Usuario creado: {user.username}. Envio correo: {detail}')
            if sent:
                flash('Usuario creado correctamente. La clave temporal fue enviada al correo registrado.', 'success')
            else:
                flash(f'Usuario creado, pero no se pudo enviar el correo. Detalle SMTP: {detail}. Revisa la configuración antes de entregar acceso.', 'warning')
            return redirect(url_for('admin_users'))
        users = User.query.order_by(User.created_at.desc()).all()
        return render_template('admin_users.html', users=users)


    @app.route('/admin/users/<int:user_id>/resend-access', methods=['POST'])
    @login_required
    def admin_user_resend_access(user_id: int):
        require_user_admin()
        user = User.query.get_or_404(user_id)
        if not user.email:
            flash('El usuario no tiene correo registrado para enviar el acceso.', 'warning')
            return redirect(url_for('admin_users'))
        temp_password = generate_temporary_password()
        user.set_password(temp_password)
        user.must_change_password = True
        db.session.commit()
        sent, detail = send_welcome_email(user, temp_password, purpose='reset')
        log_action('resend_access', 'User', user.id, f'Reenvio de acceso temporal a {user.email}: {detail}')
        if sent:
            flash('Nueva clave temporal enviada al correo del usuario.', 'success')
        else:
            flash(f'No se pudo enviar la nueva clave temporal. Detalle SMTP: {detail}', 'danger')
        return redirect(url_for('admin_users'))

    @app.route('/admin/users/<int:user_id>/edit', methods=['GET', 'POST'])
    @login_required
    def admin_user_edit(user_id: int):
        require_user_admin()
        user = User.query.get_or_404(user_id)
        if request.method == 'POST':
            email = request.form.get('email', '').strip().lower()
            if not email:
                flash('El correo electrónico es obligatorio.', 'danger')
                return render_template('admin_user_edit.html', user=user)
            duplicate = User.query.filter(
                User.id != user.id,
                or_(db.func.lower(User.username) == email, db.func.lower(User.email) == email)
            ).first()
            if duplicate:
                flash('Ya existe otro usuario registrado con ese correo. No se puede duplicar.', 'warning')
                return render_template('admin_user_edit.html', user=user)
            user.email = email
            user.username = email
            user.full_name = request.form.get('full_name', '').strip() or user.username
            user.role = request.form.get('role', user.role)
            user.is_active = bool(request.form.get('is_active'))
            user.can_create_projects = bool(request.form.get('can_create_projects'))
            user.can_join_projects = bool(request.form.get('can_join_projects'))
            user.can_manage_records = bool(request.form.get('can_manage_records'))
            user.can_manage_users = bool(request.form.get('can_manage_users'))
            user.can_view_audit = bool(request.form.get('can_view_audit'))
            password = request.form.get('password', '').strip()
            if password:
                user.set_password(password)
                user.must_change_password = True
            db.session.commit()
            log_action('update', 'User', user.id, f'Usuario actualizado: {user.username}')
            flash('Usuario actualizado.', 'success')
            return redirect(url_for('admin_users'))
        return render_template('admin_user_edit.html', user=user)

    @app.route('/admin/audit')
    @login_required
    def admin_audit():
        if not (g.current_user.is_admin or g.current_user.can_view_audit):
            abort(403)
        logs = AuditLog.query.order_by(AuditLog.created_at.desc()).limit(300).all()
        return render_template('admin_audit.html', logs=logs)

    @app.cli.command('init-db')
    def init_db_command():
        db.create_all()
        ensure_schema_updates()
        print('Base de datos inicializada.')

    with app.app_context():
        db.create_all()
        ensure_schema_updates()

    return app


def ensure_schema_updates() -> None:
    inspector = inspect(db.engine)
    tables = set(inspector.get_table_names())

    def add_columns(table: str, definitions: dict[str, str]) -> None:
        if table not in tables:
            return
        columns = {col['name'] for col in inspector.get_columns(table)}
        for column_name, ddl in definitions.items():
            if column_name not in columns:
                db.session.execute(text(f"ALTER TABLE {table} ADD COLUMN {ddl}"))

    add_columns('user', {
        'must_change_password': 'must_change_password BOOLEAN DEFAULT 0',
    })
    add_columns('project', {
        'updated_at': 'updated_at DATETIME',
        'created_by_id': 'created_by_id INTEGER',
    })
    add_columns('test_record', {
        'validation_notes': 'validation_notes TEXT',
        'resolved_at': 'resolved_at DATETIME',
        'created_by_id': 'created_by_id INTEGER',
        'updated_by_id': 'updated_by_id INTEGER',
    })
    add_columns('evidence', {
        'uploaded_by_id': 'uploaded_by_id INTEGER',
    })
    if 'test_record' in tables:
        db.session.execute(text("UPDATE test_record SET status = 'Subsanada' WHERE status IN ('Resuelta', 'Ajuste aplicado')"))
        db.session.execute(text("UPDATE test_record SET status = 'Reportada' WHERE status IN ('En análisis', 'En ajuste proveedor', 'Se mantiene')"))
    ensure_default_admin()
    db.session.commit()


def ensure_default_admin() -> None:
    """Crea o migra el administrador principal sin sobrescribir su contraseña.

    La contraseña solo se define cuando el usuario no existe o cuando la cuenta
    heredada no tiene hash. Después de eso, cualquier cambio realizado desde la
    aplicación queda persistido en la base de datos.
    """
    admin_email = os.getenv('ADMIN_EMAIL', 'cparra@coopetrol.coop').strip().lower()
    initial_password = 'cparra123'

    # Normaliza campos base.
    for user in User.query.order_by(User.id.asc()).all():
        user.email = (user.email or user.username or '').strip().lower()
        user.username = (user.username or user.email or '').strip().lower()

    db.session.flush()

    candidates = User.query.filter(
        or_(
            db.func.lower(User.username) == admin_email,
            db.func.lower(User.email) == admin_email,
            db.func.lower(User.username) == 'admin@local',
            db.func.lower(User.email) == 'admin@local',
        )
    ).order_by(User.id.asc()).all()

    admin = None
    for user in candidates:
        if (user.username or '').lower() == admin_email or (user.email or '').lower() == admin_email:
            admin = user
            break
    if admin is None and candidates:
        admin = candidates[0]

    created_now = False
    if admin is None:
        admin = User(
            username=admin_email,
            full_name='Administrador',
            email=admin_email,
            role='admin',
            must_change_password=True,
            **ROLE_DEFAULTS['admin'],
        )
        admin.set_password(initial_password)
        created_now = True
        db.session.add(admin)
        db.session.flush()
        db.session.add(AuditLog(
            user_id=admin.id,
            action='bootstrap',
            entity_type='User',
            entity_id=admin.id,
            detail='Usuario administrador inicial creado',
        ))

    # Desactiva y renombra cualquier cuenta genérica o duplicada para que no haya conflictos UNIQUE.
    for user in User.query.order_by(User.id.asc()).all():
        if user.id == admin.id:
            continue
        is_legacy_admin = (user.username or '').lower() == 'admin@local' or (user.email or '').lower() == 'admin@local'
        is_duplicate_admin_email = (user.username or '').lower() == admin_email or (user.email or '').lower() == admin_email
        if is_legacy_admin or is_duplicate_admin_email:
            user.is_active = False
            user.username = f'desactivado_{user.id}_{user.username or "usuario"}'[:80]
            user.email = f'desactivado_{user.id}_{user.email or "usuario@local"}'[:180]
            user.role = 'viewer'
            user.can_create_projects = False
            user.can_join_projects = False
            user.can_manage_records = False
            user.can_manage_users = False
            user.can_view_audit = False

    db.session.flush()

    admin.username = admin_email
    admin.email = admin_email
    admin.full_name = admin.full_name or 'Administrador'
    admin.role = 'admin'
    admin.is_active = True
    admin.can_create_projects = True
    admin.can_join_projects = True
    admin.can_manage_records = True
    admin.can_manage_users = True
    admin.can_view_audit = True

    # No sobrescribir la contraseña si ya existe. Esto evita que al reiniciar
    # la app vuelva a quedar la clave inicial.
    if not getattr(admin, 'password_hash', None):
        admin.set_password(initial_password)
        admin.must_change_password = True
    elif created_now:
        admin.must_change_password = True


def login_required(view):
    from functools import wraps
    @wraps(view)
    def wrapper(*args, **kwargs):
        if not g.get('current_user'):
            return redirect(url_for('login', next=request.path))
        return view(*args, **kwargs)
    return wrapper


def log_action(action: str, entity_type: str, entity_id: int | None = None, detail: str | None = None) -> None:
    try:
        db.session.add(AuditLog(
            user_id=g.current_user.id if getattr(g, 'current_user', None) else None,
            action=action,
            entity_type=entity_type,
            entity_id=entity_id,
            detail=detail,
            ip_address=request.remote_addr if request else None,
        ))
        db.session.commit()
    except Exception:
        db.session.rollback()


def is_project_member(project_id: int, user_id: int) -> bool:
    return ProjectMember.query.filter_by(project_id=project_id, user_id=user_id).first() is not None


def visible_projects_query():
    user = g.current_user
    if user.is_admin or user.can_create_projects or user.can_join_projects:
        return Project.query
    return Project.query.join(ProjectMember).filter(ProjectMember.user_id == user.id)


def visible_records_query():
    user = g.current_user
    if user.is_admin:
        return TestRecord.query
    return TestRecord.query.join(ProjectMember, ProjectMember.project_id == TestRecord.project_id).filter(ProjectMember.user_id == user.id)


def recordable_projects_query():
    user = g.current_user
    if user.is_admin:
        return Project.query
    if not user.can_manage_records:
        return Project.query.filter(db.text('1 = 0'))
    return Project.query.join(ProjectMember).filter(
        ProjectMember.user_id == user.id,
        ProjectMember.can_register_records.is_(True),
    )


def is_record_closed(record: TestRecord | None) -> bool:
    return bool(record and record.status in CLOSED_STATUSES)


def can_user_create_record_in_project(project: Project | None, user: User | None = None) -> bool:
    user = user or g.get('current_user')
    if not project or not user:
        return False
    if user.is_admin:
        return True
    membership = ProjectMember.query.filter_by(project_id=project.id, user_id=user.id).first()
    return bool(user.can_manage_records and membership and membership.can_register_records)


def can_user_manage_record(record: TestRecord | None, user: User | None = None) -> bool:
    user = user or g.get('current_user')
    if not record or not user or is_record_closed(record):
        return False
    return can_user_create_record_in_project(record.project, user)


def require_project_access(project: Project) -> None:
    if g.current_user.is_admin or g.current_user.can_create_projects:
        return
    if is_project_member(project.id, g.current_user.id):
        return
    abort(403)


def require_record_permission(project: Project | None) -> None:
    if not project:
        abort(404)
    user = g.current_user
    if user.is_admin:
        return
    membership = ProjectMember.query.filter_by(project_id=project.id, user_id=user.id).first()
    if user.can_manage_records and membership and membership.can_register_records:
        return
    flash('No tienes permiso para registrar o editar novedades en este proyecto. Debes tener el permiso “Registrar/editar incidencias” y estar asociado al proyecto con autorización para registrar.', 'warning')
    abort(403)


def is_project_creator(project: Project, user: User | None = None) -> bool:
    user = user or g.get('current_user')
    return bool(project and user and project.created_by_id == user.id)


def is_project_role_admin(project: Project, user: User | None = None) -> bool:
    user = user or g.get('current_user')
    if not project or not user:
        return False
    membership = ProjectMember.query.filter_by(project_id=project.id, user_id=user.id).first()
    return bool(membership and membership.project_role == 'Administrador del proyecto')


def can_user_edit_project_details(project: Project, user: User | None = None) -> bool:
    # El nombre y la descripción solo los puede modificar el usuario que creó el proyecto.
    return is_project_creator(project, user)


def can_user_manage_project_members(project: Project, user: User | None = None) -> bool:
    user = user or g.get('current_user')
    if not project or not user:
        return False
    return bool(user.is_admin or is_project_creator(project, user) or is_project_role_admin(project, user))


def can_user_manage_privileged_project_members(project: Project, user: User | None = None) -> bool:
    user = user or g.get('current_user')
    return bool(user and (user.is_admin or is_project_creator(project, user)))


def can_user_assign_project_role(project: Project, target_user: User, new_role: str, user: User | None = None) -> bool:
    user = user or g.get('current_user')
    if not can_user_manage_project_members(project, user):
        return False
    if can_user_manage_privileged_project_members(project, user):
        return True
    existing = ProjectMember.query.filter_by(project_id=project.id, user_id=target_user.id).first()
    target_is_privileged = bool(
        target_user.is_admin
        or project.created_by_id == target_user.id
        or (existing and existing.project_role == 'Administrador del proyecto')
    )
    # Un administrador de proyecto no creador solo puede administrar perfiles no privilegiados.
    return bool(new_role != 'Administrador del proyecto' and not target_is_privileged)


def can_user_remove_project_member(project: Project, membership: ProjectMember, user: User | None = None) -> bool:
    user = user or g.get('current_user')
    if not can_user_manage_project_members(project, user):
        return False
    if can_user_manage_privileged_project_members(project, user):
        return True
    target_user = membership.user
    target_is_privileged = bool(
        target_user.is_admin
        or project.created_by_id == target_user.id
        or membership.project_role == 'Administrador del proyecto'
    )
    return not target_is_privileged


def require_project_owner(project: Project) -> None:
    if can_user_edit_project_details(project):
        return
    abort(403)


def require_project_member_admin(project: Project) -> None:
    if can_user_manage_project_members(project):
        return
    abort(403)


def require_project_admin(project: Project) -> None:
    require_project_member_admin(project)


def require_user_admin() -> None:
    if not (g.current_user.is_admin or g.current_user.can_manage_users):
        abort(403)

def handle_record_form(project: Project | None = None, projects=None):
    projects = projects or Project.query.order_by(Project.name.asc()).all()

    if request.method == 'POST':
        record = TestRecord()
        update_record_from_form(record)
        if not record.project_id:
            flash('Selecciona el proyecto.', 'danger')
            selected_project = Project.query.get(record.project_id) if record.project_id else project
            return render_template('record_form.html', projects=projects, record=record, selected_project=selected_project)
        if not record.module_name or not record.environment or not record.browser or not record.expected_result or not record.actual_result or not record.responsible or not record.execution_date:
            flash('Todos los campos de texto y la fecha de ejecución son obligatorios. Las evidencias son opcionales.', 'danger')
            selected_project = Project.query.get(record.project_id) if record.project_id else project
            return render_template('record_form.html', projects=projects, record=record, selected_project=selected_project)

        require_record_permission(Project.query.get(record.project_id))
        record.created_by_id = g.current_user.id
        record.updated_by_id = g.current_user.id
        db.session.add(record)
        db.session.commit()
        log_action('create', 'TestRecord', record.id, f'Novedad creada en proyecto #{record.project_id}')
        save_uploaded_files(request.files.getlist('evidences'), record.id)
        flash('Novedad de prueba registrada con éxito.', 'success')
        return redirect(url_for('record_detail', record_id=record.id))

    return render_template('record_form.html', projects=projects, record=None, selected_project=project)


def update_record_from_form(record: TestRecord) -> None:
    execution_date_raw = request.form.get('execution_date', '').strip()
    record.execution_date = datetime.strptime(execution_date_raw, '%Y-%m-%dT%H:%M') if execution_date_raw else None
    record.module_name = request.form.get('module_name', '').strip()
    record.environment = request.form.get('environment', '').strip()
    record.browser = request.form.get('browser', '').strip()
    # Campos retirados del formulario. Se conservan en el modelo para compatibilidad
    # con bases existentes, pero se diligencian automáticamente.
    project_id_raw = request.form.get('project_id')
    project_name = ''
    if project_id_raw:
        project_obj = Project.query.get(int(project_id_raw))
        project_name = project_obj.name if project_obj else ''
    record.application_name = project_name or 'No aplica'
    record.version = ''
    record.platform = ''
    record.test_type = request.form.get('test_type', 'Funcional')
    record.status = request.form.get('status', 'Reportada')
    record.severity = request.form.get('severity', 'Media')
    record.priority = request.form.get('priority', 'Media')
    record.expected_result = request.form.get('expected_result', '').strip()
    record.actual_result = request.form.get('actual_result', '').strip()
    if record.id:
        record.validation_notes = request.form.get('validation_notes', '').strip()
    else:
        record.validation_notes = ''
    generated_title = record.actual_result or record.expected_result or record.module_name or 'Incidencia sin descripción'
    record.title = generated_title[:180]
    record.responsible = request.form.get('responsible', '').strip()
    record.project_id = request.form.get('project_id', type=int)
    if record.status in CLOSED_STATUSES and not record.resolved_at:
        record.resolved_at = datetime.utcnow()
    elif record.status not in CLOSED_STATUSES:
        record.resolved_at = None


def safe_project_name(project: Project) -> str:
    return secure_filename(project.name.lower().replace(' ', '_')) or f'proyecto_{project.id}'


def summarize_for_vendor(record: TestRecord) -> str:
    """Build a concise, readable summary for CSV/XLSX vendor exports."""
    sections = []
    if record.actual_result:
        sections.append(f"Hallazgo: {record.actual_result.strip()}")
    if record.expected_result:
        sections.append(f"Resultado esperado: {record.expected_result.strip()}")
    if not sections:
        sections.append(record.title or 'Sin descripcion')
    return "\n".join(sections)


def vendor_rows(project: Project):
    rows = TestRecord.query.filter_by(project_id=project.id).order_by(TestRecord.created_at.asc()).all()
    data = []
    for item in rows:
        data.append([
            item.id,
            project.name,
            item.title,
            item.module_name or '',
            item.environment or '',
            item.browser or '',
            item.test_type,
            item.status,
            item.severity,
            item.priority,
            summarize_for_vendor(item),
            item.responsible or '',
            item.execution_date.strftime('%Y-%m-%d %H:%M') if item.execution_date else '',
            item.validation_notes or '',
            item.resolved_at.strftime('%Y-%m-%d %H:%M') if item.resolved_at else '',
        ])
    return data


def build_project_csv(project: Project) -> io.BytesIO:
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        'ID', 'Proyecto', 'Novedad', 'Modulo', 'Ambiente',
        'Navegador/Dispositivo', 'Tipo de prueba', 'Estado',
        'Severidad', 'Prioridad', 'Descripcion resumida', 'Responsable', 'Fecha de ejecucion',
        'Validacion QA', 'Fecha de cierre'
    ])
    for row in vendor_rows(project):
        writer.writerow(row)

    mem = io.BytesIO()
    mem.write(output.getvalue().encode('utf-8-sig'))
    mem.seek(0)
    return mem


def build_project_xlsx(project: Project) -> io.BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = 'Novedades'
    headers = [
        'ID', 'Proyecto', 'Novedad', 'Modulo', 'Ambiente',
        'Navegador/Dispositivo', 'Tipo de prueba', 'Estado',
        'Severidad', 'Prioridad', 'Descripcion resumida', 'Responsable', 'Fecha de ejecucion',
        'Validacion QA', 'Fecha de cierre'
    ]
    ws.append(headers)
    for row in vendor_rows(project):
        ws.append(row)

    header_fill = PatternFill('solid', fgColor='1F4E78')
    header_font = Font(color='FFFFFF', bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    widths = [8, 20, 30, 18, 16, 22, 16, 18, 12, 12, 45, 18, 18, 24, 18]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical='top', wrap_text=True)
    ws.freeze_panes = 'A2'

    mem = io.BytesIO()
    wb.save(mem)
    mem.seek(0)
    return mem




def qa_pdf_status_color(status):
    if status == 'Subsanada':
        return colors.HexColor('#198754')
    if status == 'Descartada':
        return colors.HexColor('#6c757d')
    return colors.HexColor('#fd7e14')

def build_project_pdf(project: Project) -> io.BytesIO:
    """Genera el PDF corporativo usando el membrete oficial como base visual."""
    mem = io.BytesIO()
    doc = SimpleDocTemplate(
        mem,
        pagesize=A4,
        rightMargin=1.85 * cm,
        leftMargin=2.35 * cm,
        topMargin=3.85 * cm,
        bottomMargin=2.85 * cm,
    )

    coop_green = colors.HexColor('#00843D')
    coop_yellow = colors.HexColor('#F7D117')
    coop_dark = colors.HexColor('#1D1D1B')
    border_green = colors.HexColor('#00843D')
    light_green = colors.HexColor('#EAF6EF')
    light_yellow = colors.HexColor('#FFF8CC')
    soft_gray = colors.HexColor('#F7F7F4')
    white = colors.HexColor('#FFFFFF')
    mid_gray = colors.HexColor('#D9D9D4')
    text_color = colors.HexColor('#202020')
    width, height = A4
    content_width = width - doc.leftMargin - doc.rightMargin
    logo_path = BASE_DIR / 'static' / 'branding' / 'logo_coopetrol.png'
    supersolidario_path = BASE_DIR / 'static' / 'branding' / 'supersolidario.png'

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CoopTitle',
        parent=styles['Title'],
        fontName='Helvetica-Bold',
        fontSize=17,
        leading=22,
        textColor=coop_dark,
        alignment=1,
        spaceAfter=5,
    )
    subtitle_style = ParagraphStyle(
        'CoopSubtitle',
        parent=styles['BodyText'],
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#555555'),
        alignment=1,
        spaceAfter=10,
    )
    intro_style = ParagraphStyle(
        'CoopIntro',
        parent=styles['BodyText'],
        fontSize=8.8,
        leading=12,
        textColor=text_color,
        spaceAfter=3,
    )
    section_title_style = ParagraphStyle(
        'CoopSectionTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=13.5,
        leading=17,
        textColor=coop_green,
        spaceBefore=8,
        spaceAfter=6,
    )
    label_style = ParagraphStyle(
        'CoopLabel',
        parent=styles['BodyText'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=12,
        textColor=coop_green,
        spaceBefore=5,
        spaceAfter=3,
    )
    body_style = ParagraphStyle(
        'CoopBody',
        parent=styles['BodyText'],
        fontSize=9.2,
        leading=13,
        textColor=text_color,
        spaceAfter=4,
    )
    small_style = ParagraphStyle(
        'CoopSmall',
        parent=styles['BodyText'],
        fontSize=7.8,
        leading=9.2,
        textColor=colors.HexColor('#FFFFFF'),
    )

    records = TestRecord.query.filter_by(project_id=project.id).order_by(TestRecord.created_at.asc()).all()
    pending_records = [r for r in records if r.status not in CLOSED_STATUSES]
    total_count = len(records)
    pending_count = len(pending_records)
    subsanadas_count = sum(1 for r in records if r.status == 'Subsanada')
    descartadas_count = sum(1 for r in records if r.status == 'Descartada')

    def draw_membrete(canvas, doc_obj):
        """Dibuja el membrete corporativo siguiendo el formato FR-CD-23.

        Ajustes solicitados:
        - Logo Coopetrol centrado en la parte superior.
        - Logo Supersolidaria lateral izquierdo, más visible.
        - Sin texto "FR-CD-23 Hoja Membrete".
        - Sin línea azul; solo acentos verde y amarillo corporativos.
        - Pie de página centrado, sin fondo de color.
        """
        canvas.saveState()

        # Marco exterior del documento, como hoja membretada.
        canvas.setStrokeColor(colors.HexColor('#D7E0DB'))
        canvas.setLineWidth(0.85)
        canvas.rect(1.2 * cm, 0.92 * cm, width - 2.4 * cm, height - 1.75 * cm, fill=0, stroke=1)

        # Logo principal centrado.
        if logo_path.exists():
            logo_w = 8.2 * cm
            logo_h = 2.15 * cm
            canvas.drawImage(
                str(logo_path),
                (width - logo_w) / 2,
                height - 3.0 * cm,
                width=logo_w,
                height=logo_h,
                preserveAspectRatio=True,
                mask='auto',
            )

        # Logo Supersolidaria en lateral izquierdo, más grande.
        if supersolidario_path.exists():
            side_w = 0.86 * cm
            side_h = 5.4 * cm
            canvas.drawImage(
                str(supersolidario_path),
                1.34 * cm,
                2.15 * cm,
                width=side_w,
                height=side_h,
                preserveAspectRatio=True,
                mask='auto',
            )

        # Línea corporativa: verde + amarillo, sin azul.
        line_y = height - 3.42 * cm
        line_x = 1.2 * cm
        line_w = width - 2.4 * cm
        canvas.setLineWidth(0.5)
        canvas.setStrokeColor(coop_green)
        canvas.line(line_x, line_y, line_x + line_w * 0.72, line_y)
        canvas.setStrokeColor(coop_yellow)
        canvas.line(line_x + line_w * 0.72, line_y, line_x + line_w, line_y)

        # Pie de página institucional centrado, sin fondo.
        footer_lines = [
            'Dirección General: Cra13A # 34-72 Oficina 301 - Bogotá D.C., Colombia',
            'Agencias: Barrancabermeja, Barranquilla, Bogotá Teusaquillo, Bogotá Norte, Bucaramanga, Cali, Cartagena, Cúcuta, La Dorada, Manizales, Medellín,',
            'Neiva, Orito, Pasto, Tibú y Villavicencio.',
            'Puntos de Atención: Ibagué, Valledupar, Mamonal y Tumaco.',
            'Servicio al cliente: www.coopetrol.coop - Infocoopetrol@coopetrl.coop - PBX: 6014918690 linea gratuita nacional 01 8000 423761',
        ]
        canvas.setFont('Helvetica', 4.8)
        canvas.setFillColor(colors.HexColor('#1D1D1B'))
        footer_y = 1.58 * cm
        for line in footer_lines:
            canvas.drawCentredString(width / 2, footer_y, line)
            footer_y -= 0.17 * cm

        canvas.setFont('Helvetica', 6.4)
        canvas.setFillColor(colors.HexColor('#4A4A4A'))
        canvas.drawRightString(width - 1.28 * cm, 0.70 * cm, f'Página {doc_obj.page}')
        canvas.restoreState()

    story = []
    story.append(Paragraph(f'Informe de Pruebas - {pdf_text(project.name)}', title_style))
    story.append(Paragraph(f'Generado el {datetime.now().strftime("%Y-%m-%d %H:%M")}', subtitle_style))
    # story.append(coop_rule(content_width, coop_green, coop_yellow))
    story.append(Spacer(1, 0.18 * cm))

    metrics = Table([
        [
            corporate_metric('Total incidencias', str(total_count), coop_green),
            corporate_metric('Pendientes', str(pending_count), coop_yellow),
            corporate_metric('Subsanadas', f'{subsanadas_count} / {total_count}', coop_green),
            corporate_metric('Descartadas', str(descartadas_count), coop_dark),
        ]
    ], colWidths=[content_width / 4.0] * 4)
    metrics.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(metrics)
    story.append(Spacer(1, 0.25 * cm))
    story.append(corporate_note(
        'Este informe relaciona únicamente las incidencias pendientes de subsanación. Las novedades marcadas como subsanadas o descartadas no se incluyen en el detalle.',
        intro_style,
        content_width,
        light_green,
        coop_green,
    ))
    story.append(Spacer(1, 0.18 * cm))

    if not pending_records:
        story.append(corporate_note('No hay incidencias pendientes por reportar.', body_style, content_width, light_yellow, coop_yellow))
    else:
        for index, record in enumerate(pending_records, start=1):
            if index > 1:
                story.append(PageBreak())
            story.append(Paragraph(f'Incidencia {index}', section_title_style))
            # story.append(coop_rule(content_width, 0, coop_yellow, height_cm=0.0035))
            story.append(Spacer(1, 0.12 * cm))

            meta = Table([
                [
                    meta_item('ID', f'#{record.id}', small_style),
                    meta_item('Estado', record.status, small_style),
                    meta_item('Severidad', record.severity, small_style),
                    meta_item('Prioridad', record.priority, small_style),
                ],
                [
                    meta_item('Módulo', record.module_name or 'No registra', small_style),
                    meta_item('Ambiente', record.environment or 'No registra', small_style),
                    meta_item('Tipo prueba', record.test_type or 'No registra', small_style),
                    meta_item('Fecha', record.execution_date.strftime('%Y-%m-%d %H:%M') if record.execution_date else 'No registra', small_style),
                ],
            ], colWidths=[content_width / 4.0] * 4)
            meta.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), white),
                ('BOX', (0, 0), (-1, -1), 0.65, mid_gray),
                ('INNERGRID', (0, 0), (-1, -1), 0.35, mid_gray),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 7),
                ('RIGHTPADDING', (0, 0), (-1, -1), 7),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ]))
            story.append(meta)
            story.append(Spacer(1, 0.18 * cm))

            story.append(corporate_text_block('Descripción del hallazgo', record.actual_result or 'No registra descripción del hallazgo.', body_style, label_style, content_width, '#FFFFFF', '#007934'))
            story.append(Spacer(1, 0.08 * cm))
            story.append(corporate_text_block('Resultado esperado', record.expected_result or 'No registra resultado esperado.', body_style, label_style, content_width, '#FFFFFF', '#1c1b1a'))
            if record.validation_notes:
                story.append(Spacer(1, 0.08 * cm))
                story.append(corporate_text_block('Validación QA', record.validation_notes, body_style, label_style, content_width, '#F7F7F4', '#00843D'))

            image_exts = {'jpg', 'jpeg', 'png'}
            image_paths = [UPLOAD_DIR / ev.file_name for ev in record.evidences if (ev.file_type or '').lower() in image_exts and (UPLOAD_DIR / ev.file_name).exists()]
            if image_paths:
                story.append(Spacer(1, 0.16 * cm))
                story.append(Paragraph('Evidencia:', label_style))
                # En el formato membreteado las evidencias se muestran en paginas dedicadas
                # para no invadir el pie de pagina corporativo y conservar legibilidad.
                # story.append(PageBreak())
                story.extend(build_corporate_evidence_blocks(image_paths[:10], content_width, coop_green, mid_gray))

            other_evidences = [ev.original_name for ev in record.evidences if (ev.file_type or '').lower() not in image_exts]
            if other_evidences:
                story.append(Spacer(1, 0.08 * cm))
                story.append(corporate_text_block('Archivos asociados', '<br/>'.join(pdf_text(name) for name in other_evidences[:10]), body_style, label_style, content_width, '#FFFFFF', '#D9D9D4', escape_value=False))

    doc.build(story, onFirstPage=draw_membrete, onLaterPages=draw_membrete)
    mem.seek(0)
    return mem


def corporate_metric(label: str, value: str, accent_color):
    label_style = ParagraphStyle(
        'CorpMetricLabel_' + label.replace(' ', '_'),
        fontName='Helvetica-Bold',
        fontSize=7.5,
        leading=9,
        textColor=colors.HexColor('#444444'),
        alignment=1,
    )
    value_style = ParagraphStyle(
        'CorpMetricValue_' + label.replace(' ', '_'),
        fontName='Helvetica-Bold',
        fontSize=15,
        leading=17,
        textColor=colors.HexColor('#1D1D1B'),
        alignment=1,
    )
    card = Table([
        [''],
        [Paragraph(pdf_text(label), label_style)],
        [Paragraph(pdf_text(value), value_style)],
    ], colWidths=[3.8 * cm], rowHeights=[0.12 * cm, None, None])
    card.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), accent_color),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('BOX', (0, 0), (-1, -1), 0.55, colors.HexColor('#CFCFC8')),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 1), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 5),
    ]))
    return card


def coop_rule(width, green, yellow, height_cm=0.05):
    w1 = width * 0.82
    w2 = width - w1
    rule = Table([['', '']], colWidths=[w1, w2], rowHeights=[height_cm * cm])
    rule.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), green),
        ('BACKGROUND', (1, 0), (1, 0), yellow),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    return rule


def corporate_note(text: str, style: ParagraphStyle, width, background, border):
    table = Table([[Paragraph(pdf_text(text), style)]], colWidths=[width])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), background),
        ('BOX', (0, 0), (-1, -1), 0.55, border),
        ('LEFTPADDING', (0, 0), (-1, -1), 9),
        ('RIGHTPADDING', (0, 0), (-1, -1), 9),
        ('TOPPADDING', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
    ]))
    return table


def corporate_text_block(title: str, value: str, body_style: ParagraphStyle, label_style: ParagraphStyle, width, background_hex: str, border_hex: str, escape_value: bool = True):
    content = pdf_text(value) if escape_value else value
    title_style = ParagraphStyle(
        'CorporateBlockTitle_' + title.replace(' ', '_'),
        parent=label_style,
        fontName='Helvetica-Bold',
        fontSize=9.3,
        leading=12,
        textColor=colors.HexColor(border_hex),
        spaceBefore=0,
        spaceAfter=3,
    )
    table = Table([
        [Paragraph(pdf_text(title), title_style)],
        [Paragraph(content, body_style)],
    ], colWidths=[width])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(background_hex)),
        ('BOX', (0, 0), (-1, -1), 0.55, colors.HexColor(border_hex)),
        ('LEFTPADDING', (0, 0), (-1, -1), 9),
        ('RIGHTPADDING', (0, 0), (-1, -1), 9),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
    ]))
    return table


def meta_item(label: str, value: str, style: ParagraphStyle):
    return Paragraph(f'<font color="#00843D"><b>{pdf_text(label)}</b></font><br/><font color="#1D1D1B"><b>{pdf_text(value)}</b></font>', style)


def build_corporate_evidence_blocks(image_paths: list[Path], content_width, coop_green, border_color):
    blocks = []
    for idx, image_path in enumerate(image_paths, start=1):
        blocks.append(build_corporate_evidence_cell(image_path, idx, content_width, coop_green, border_color))
        blocks.append(Spacer(1, 0.22 * cm))
    return blocks


def build_corporate_evidence_cell(image_path: Path, index: int, content_width, coop_green, border_color):
    label_style = ParagraphStyle(
        'CorporateEvidenceLabel_' + str(index),
        fontName='Helvetica-Bold',
        fontSize=9.2,
        leading=12,
        textColor=coop_green,
        spaceAfter=5,
    )
    try:
        img = ImageReader(str(image_path))
        width, height = img.getSize()
        max_width = content_width - 0.7 * cm
        max_height = 10.2 * cm
        scale = min(max_width / width, max_height / height, 1)
        reportlab_image = RLImage(str(image_path), width=width * scale, height=height * scale)
        reportlab_image.hAlign = 'CENTER'
        inner = Table([
            [Paragraph(f'Evidencia {index}', label_style)],
            [reportlab_image],
        ], colWidths=[content_width - 0.4 * cm])
        inner.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.white),
            ('BOX', (0, 0), (-1, -1), 0.45, border_color),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
            ('RIGHTPADDING', (0, 0), (-1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 7),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('ALIGN', (0, 1), (0, 1), 'CENTER'),
        ]))
        return KeepTogether([inner])
    except Exception:
        file_style = ParagraphStyle('CorporateEvidenceFileName', fontSize=8.5, textColor=colors.HexColor('#555555'))
        return KeepTogether([Paragraph(f'Evidencia {index}', label_style), Paragraph(pdf_text(image_path.name), file_style)])


def pdf_text(value: str) -> str:
    safe = (value or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return safe.replace('\n', '<br/>')

def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def save_uploaded_files(files, record_id: int) -> None:
    for file in files:
        if not file or not file.filename:
            continue
        if not allowed_file(file.filename):
            continue
        original_name = secure_filename(file.filename)
        ext = original_name.rsplit('.', 1)[1].lower()
        unique_name = f"{uuid.uuid4().hex}.{ext}"
        file_path = UPLOAD_DIR / unique_name
        file.save(file_path)
        db.session.add(Evidence(
            original_name=original_name,
            file_name=unique_name,
            file_type=ext,
            uploaded_by_id=g.current_user.id if getattr(g, 'current_user', None) else None,
            test_record_id=record_id,
        ))
    db.session.commit()


def generate_temporary_password(length: int = 14) -> str:
    alphabet = string.ascii_letters + string.digits + "!@#$%&*"
    while True:
        password = ''.join(secrets.choice(alphabet) for _ in range(length))
        if (any(c.islower() for c in password) and any(c.isupper() for c in password)
                and any(c.isdigit() for c in password) and any(c in "!@#$%&*" for c in password)):
            return password


def send_welcome_email(user: User, temporary_password: str, purpose: str = 'welcome') -> tuple[bool, str]:
    smtp_host = os.getenv('SMTP_HOST', '').strip()
    smtp_port = int(os.getenv('SMTP_PORT', '587'))
    smtp_user = os.getenv('SMTP_USER', '').strip()
    smtp_password = os.getenv('SMTP_PASSWORD', '')
    smtp_from = os.getenv('SMTP_FROM', smtp_user).strip()
    smtp_tls = os.getenv('SMTP_USE_TLS', os.getenv('SMTP_TLS', 'true')).lower() not in {'0', 'false', 'no'}
    smtp_ssl = os.getenv('SMTP_USE_SSL', os.getenv('SMTP_SSL', 'false')).lower() in {'1', 'true', 'yes'}
    if not smtp_host or not smtp_from or not user.email:
        return False, 'faltan SMTP_HOST, SMTP_FROM o correo del usuario'

    msg = EmailMessage()
    msg['Subject'] = 'Nueva clave temporal - QA Novedades' if purpose == 'reset' else 'Acceso a QA Novedades'
    msg['From'] = smtp_from
    msg['To'] = user.email
    body = (
        f"Hola {user.display_name},\n\n"
        + ("Se generó una nueva clave temporal para tu usuario en QA Novedades.\n\n" if purpose == 'reset' else "Se creó tu usuario en QA Novedades.\n\n")
        + f"Usuario: {user.email}\n"
        f"Clave temporal: {temporary_password}\n\n"
        "Al ingresar, la aplicación solicitará cambiar esta clave inmediatamente.\n\n"
        "Si no esperabas este acceso, informa al administrador de la plataforma.\n"
    )
    msg.set_content(body)
    try:
        smtp_class = smtplib.SMTP_SSL if smtp_ssl else smtplib.SMTP
        with smtp_class(smtp_host, smtp_port, timeout=20) as server:
            if smtp_tls and not smtp_ssl:
                server.starttls()
            if smtp_user:
                server.login(smtp_user, smtp_password)
            server.send_message(msg)
        return True, f'enviado a {user.email}'
    except smtplib.SMTPAuthenticationError as exc:
        return False, f'autenticación rechazada por el servidor SMTP ({exc.smtp_code})'
    except smtplib.SMTPConnectError as exc:
        return False, f'no se pudo conectar al servidor SMTP ({exc.smtp_code})'
    except smtplib.SMTPRecipientsRefused:
        return False, 'el servidor rechazó el destinatario'
    except Exception as exc:
        return False, f'{type(exc).__name__}: {exc}'


app = create_app()


if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000, debug=False)
