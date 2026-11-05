import os, io, csv, hmac, hashlib, math
from datetime import datetime, date, timedelta
from functools import wraps
from urllib.parse import urlparse
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, abort, jsonify
from flask_wtf.csrf import CSRFProtect
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import inspect, text
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.middleware.proxy_fix import ProxyFix
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# Fail closed: never run with Flask's public/default signing key.
secret_key = os.environ.get('SECRET_KEY', '')
if (len(secret_key) < 32
        or secret_key.lower().startswith(('change-me', 'changeme', 'replace-with-'))):
    raise RuntimeError('SECRET_KEY must be a unique random value of at least 32 characters, not a placeholder.')

app.config['SECRET_KEY'] = secret_key
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
is_production = os.environ.get('APP_ENV', '').lower() == 'production'
cookie_secure = os.environ.get('COOKIE_SECURE', '1' if is_production else '0') == '1'
if is_production and not cookie_secure:
    raise RuntimeError('COOKIE_SECURE=1 is required in production.')

app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['SESSION_COOKIE_NAME'] = '__Host-forus_arvestid' if cookie_secure else 'forus_arvestid'
app.config['SESSION_COOKIE_PATH'] = '/'
app.config['SESSION_COOKIE_DOMAIN'] = None
app.config['SESSION_COOKIE_SECURE'] = cookie_secure
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)
app.config['SESSION_REFRESH_EACH_REQUEST'] = True
app.config['MAX_CONTENT_LENGTH'] = 1 * 1024 * 1024
app.config['WTF_CSRF_TIME_LIMIT'] = 60 * 60
app.config['PREFERRED_URL_SCHEME'] = 'https'

database_url = os.environ.get('DATABASE_URL', 'sqlite:///metervault.db')
if database_url.startswith('postgres://'):
    database_url = database_url.replace('postgres://', 'postgresql+psycopg://', 1)
elif database_url.startswith('postgresql://'):
    database_url = database_url.replace('postgresql://', 'postgresql+psycopg://', 1)
if is_production and not database_url.startswith('postgresql+psycopg://'):
    raise RuntimeError('Production deployment requires PostgreSQL via the psycopg v3 driver.')
app.config['SQLALCHEMY_DATABASE_URI'] = database_url

# Trust only one reverse proxy for client IP and HTTPS scheme.
# X-Forwarded-Host is deliberately not trusted.
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1)
db = SQLAlchemy(app)
CSRFProtect(app)


@app.after_request
def add_security_headers(response):
    response.headers.setdefault('X-Content-Type-Options', 'nosniff')
    response.headers.setdefault('X-Frame-Options', 'DENY')
    response.headers.setdefault('Referrer-Policy', 'strict-origin-when-cross-origin')
    response.headers.setdefault('Permissions-Policy', 'camera=(), microphone=(), geolocation=()')
    response.headers.setdefault(
        'Content-Security-Policy',
        "default-src 'self'; base-uri 'self'; form-action 'self'; "
        "frame-ancestors 'none'; object-src 'none'; img-src 'self' data:; "
        "style-src 'self' 'unsafe-inline'; script-src 'self' 'unsafe-inline'"
    )
    if request.is_secure:
        response.headers.setdefault('Strict-Transport-Security', 'max-age=31536000; includeSubDomains')
    response.headers.setdefault('Cache-Control', 'no-store, private, max-age=0')
    return response

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='user')
    active = db.Column(db.Boolean, nullable=False, default=True)
    session_version = db.Column(db.Integer, nullable=False, default=1)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    def set_password(self, password): self.password_hash = generate_password_hash(password)
    def check_password(self, password): return check_password_hash(self.password_hash, password)

class Center(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(160), nullable=False)
    active = db.Column(db.Boolean, nullable=False, default=True)
    deleted_at = db.Column(db.DateTime, nullable=True)
    sort_order = db.Column(db.Integer, nullable=False, default=0)
    categories = db.relationship('Category', backref='center', cascade='all, delete-orphan', order_by='Category.sort_order')

class Category(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    deleted_at = db.Column(db.DateTime, nullable=True)
    center_id = db.Column(db.Integer, db.ForeignKey('center.id'), nullable=False)
    name = db.Column(db.String(160), nullable=False)
    comment = db.Column(db.Text, nullable=True)
    active = db.Column(db.Boolean, nullable=False, default=True)
    sort_order = db.Column(db.Integer, nullable=False, default=0)
    meters = db.relationship('Meter', backref='category', cascade='all, delete-orphan', order_by='Meter.sort_order')

class Meter(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    category_id = db.Column(db.Integer, db.ForeignKey('category.id'), nullable=False)
    name = db.Column(db.String(160), nullable=False)
    meter_code = db.Column(db.String(100), nullable=True)
    description = db.Column(db.Text, nullable=True)
    resource_type = db.Column(db.String(40), nullable=False) # Elekter, Külm vesi, Soe vesi, Gaas, Muu
    unit = db.Column(db.String(20), nullable=False, default='m³')
    tariff_mode = db.Column(db.String(10), nullable=False, default='single')
    active = db.Column(db.Boolean, nullable=False, default=True)
    deleted_at = db.Column(db.DateTime, nullable=True)
    sort_order = db.Column(db.Integer, nullable=False, default=0)
    readings = db.relationship('Reading', backref='meter', cascade='all, delete-orphan')

class Reading(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    meter_id = db.Column(db.Integer, db.ForeignKey('meter.id'), nullable=False)
    period = db.Column(db.String(7), nullable=False) # YYYY-MM
    value = db.Column(db.Float, nullable=False)
    day_value = db.Column(db.Float, nullable=True)
    night_value = db.Column(db.Float, nullable=True)
    note = db.Column(db.String(500), nullable=True)
    entered_by = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    entered_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)
    __table_args__ = (db.UniqueConstraint('meter_id', 'period', name='uq_meter_period'),)
    user = db.relationship('User')

class AuditLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    action = db.Column(db.String(100), nullable=False)
    details = db.Column(db.String(500), nullable=False, default='')
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    user = db.relationship('User')

class LoginAttempt(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ip_hash = db.Column(db.String(64), nullable=False, index=True)
    username_hash = db.Column(db.String(64), nullable=False, index=True)
    attempted_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, index=True)


def _login_fingerprint(value):
    return hmac.new(
        app.config['SECRET_KEY'].encode('utf-8'),
        value.encode('utf-8', errors='ignore'),
        hashlib.sha256
    ).hexdigest()


def _login_is_limited(ip_hash, username_hash, now):
    cutoff = now - timedelta(minutes=15)
    per_pair = LoginAttempt.query.filter(
        LoginAttempt.ip_hash == ip_hash,
        LoginAttempt.username_hash == username_hash,
        LoginAttempt.attempted_at >= cutoff
    ).count()
    per_ip = LoginAttempt.query.filter(
        LoginAttempt.ip_hash == ip_hash,
        LoginAttempt.attempted_at >= cutoff
    ).count()
    return per_pair >= 8 or per_ip >= 30


def current_user():
    uid = session.get('user_id')
    if not uid:
        return None
    user = db.session.get(User, uid)
    if (not user or not user.active
            or session.get('session_version') != user.session_version):
        session.clear()
        return None
    return user

def login_required(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        user = current_user()
        if not user or not user.active:
            session.clear()
            return redirect(url_for('login'))
        return fn(*args, **kwargs)
    return wrapped

def admin_required(fn):
    @wraps(fn)
    @login_required
    def wrapped(*args, **kwargs):
        if current_user().role != 'admin': abort(403)
        return fn(*args, **kwargs)
    return wrapped

def log_action(action, details=''):
    db.session.add(AuditLog(user_id=current_user().id, action=action, details=details[:500]))

def valid_period(p):
    try:
        datetime.strptime(p, '%Y-%m')
        return len(p) == 7
    except (ValueError, TypeError): return False

def previous_reading(meter_id, period):
    r = Reading.query.filter(Reading.meter_id == meter_id, Reading.period < period).order_by(Reading.period.desc()).first()
    return r

@app.template_filter('num')
def format_number(value):
    if value is None: return ''
    try:
        return ('{:,.6f}'.format(float(value)).rstrip('0').rstrip('.')).replace(',', ' ')
    except (TypeError, ValueError):
        return str(value)

@app.context_processor
def inject_globals():
    return {'logged_in_user': current_user(), 'current_year': date.today().year}

@app.route('/')
@login_required
def index():
    if not current_user(): return redirect(url_for('login'))
    period = request.args.get('period', date.today().strftime('%Y-%m'))
    if not valid_period(period): period = date.today().strftime('%Y-%m')
    centers = Center.query.filter_by(active=True, deleted_at=None).order_by(Center.sort_order, Center.name).all()
    stats = {}
    for c in centers:
        meters = [
            m
            for cat in c.categories if cat.active
            for m in cat.meters
            if m.active and m.deleted_at is None
        ]
        done = Reading.query.filter(Reading.period == period, Reading.meter_id.in_([m.id for m in meters])).count() if meters else 0
        stats[c.id] = (done, len(meters))
    return render_template('index.html', centers=centers, stats=stats, period=period)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user() and current_user().active:
        return redirect(url_for('index'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()[:80]
        password = request.form.get('password', '')
        ip_hash = _login_fingerprint(request.remote_addr or 'unknown')
        username_hash = _login_fingerprint(username.casefold())
        now = datetime.utcnow()

        # Persistent throttling is shared by all Gunicorn workers and survives restarts.
        LoginAttempt.query.filter(
            LoginAttempt.attempted_at < now - timedelta(days=1)
        ).delete(synchronize_session=False)
        if _login_is_limited(ip_hash, username_hash, now):
            db.session.commit()
            flash('Sisselogimine ajutiselt piiratud. Proovi 15 minuti pärast uuesti.', 'error')
            return render_template('login.html'), 429

        user = User.query.filter_by(username=username, active=True).first()
        if user and user.check_password(password):
            session.clear()
            session.permanent = True
            session['user_id'] = user.id
            session['session_version'] = user.session_version
            db.session.commit()
            return redirect(url_for('index'))

        db.session.add(LoginAttempt(
            ip_hash=ip_hash,
            username_hash=username_hash,
            attempted_at=now
        ))
        db.session.commit()
        flash('Vale kasutajanimi või parool.', 'error')
    return render_template('login.html')


@app.route('/logout', methods=['POST'])
@login_required
def logout():
    session.clear(); return redirect(url_for('login'))

@app.route('/center/<int:center_id>')
@login_required
def center_view(center_id):
    center = db.get_or_404(Center, center_id)
    if center.deleted_at is not None or not center.active:
        abort(404)
    period = request.args.get('period', date.today().strftime('%Y-%m'))
    if not valid_period(period): period = date.today().strftime('%Y-%m')
    categories = [c for c in center.categories if c.active]
    rows = []
    total = completed = 0
    for cat in categories:
        meters = [m for m in cat.meters if m.active and m.deleted_at is None]
        for meter in meters:
            total += 1
            reading = Reading.query.filter_by(meter_id=meter.id, period=period).first()
            if reading: completed += 1
            rows.append({'meter': meter, 'category': cat, 'reading': reading, 'previous': previous_reading(meter.id, period)})
    return render_template('center.html', center=center, categories=categories, rows=rows, period=period, total=total, completed=completed)


@app.route('/category/<int:category_id>/readings', methods=['GET', 'POST'])
@login_required
def category_readings(category_id):
    category = db.get_or_404(Category, category_id)
    if category.center.deleted_at is not None or not category.center.active:
        abort(404)
    if not category.active:
        abort(404)
    period = request.values.get('period', date.today().strftime('%Y-%m'))

    if not valid_period(period):
        period = date.today().strftime('%Y-%m')

    meters = [m for m in category.meters if m.active and m.deleted_at is None]

    if request.method == 'POST':
        saved = 0
        errors = []

        for meter in meters:
            note = request.form.get(f'note_{meter.id}', '').strip()
            previous = previous_reading(meter.id, period)

            if meter.tariff_mode == 'dual':
                raw_day = request.form.get(
                    f'day_value_{meter.id}', ''
                ).strip().replace(' ', '').replace('\u00a0', '').replace(',', '.')

                raw_night = request.form.get(
                    f'night_value_{meter.id}', ''
                ).strip().replace(' ', '').replace('\u00a0', '').replace(',', '.')

                if raw_day == '' and raw_night == '':
                    continue

                if raw_day == '' or raw_night == '':
                    errors.append(
                        f'{meter.name}: sisesta nii päeva- kui ka öönäit.'
                    )
                    continue

                try:
                    day_value = float(raw_day)
                    night_value = float(raw_night)

                    if (
                        day_value < 0
                        or night_value < 0
                        or not math.isfinite(day_value)
                        or not math.isfinite(night_value)
                    ):
                        raise ValueError()

                except ValueError:
                    errors.append(
                        f'{meter.name}: kontrolli päeva- ja öönäidu väärtusi.'
                    )
                    continue

                if previous and previous.day_value is not None:
                    if day_value < previous.day_value:
                        errors.append(
                            f'{meter.name}: päevanäit on väiksem kui eelmine.'
                        )
                        continue

                if previous and previous.night_value is not None:
                    if night_value < previous.night_value:
                        errors.append(
                            f'{meter.name}: öönäit on väiksem kui eelmine.'
                        )
                        continue

                total_value = day_value + night_value

                reading = Reading.query.filter_by(
                    meter_id=meter.id,
                    period=period
                ).first()

                if reading:
                    reading.value = total_value
                    reading.day_value = day_value
                    reading.night_value = night_value
                    reading.note = note
                    reading.entered_by = current_user().id
                    reading.updated_at = datetime.utcnow()
                else:
                    reading = Reading(
                        meter_id=meter.id,
                        period=period,
                        value=total_value,
                        day_value=day_value,
                        night_value=night_value,
                        note=note,
                        entered_by=current_user().id
                    )
                    db.session.add(reading)

            else:
                raw = request.form.get(
                    f'value_{meter.id}', ''
                ).strip().replace(' ', '').replace('\u00a0', '').replace(',', '.')

                if raw == '':
                    continue

                try:
                    value = float(raw)

                    if value < 0 or not math.isfinite(value):
                        raise ValueError()

                except ValueError:
                    errors.append(
                        f'{meter.name}: sisesta korrektne näit.'
                    )
                    continue

                if previous and value < previous.value:
                    errors.append(
                        f'{meter.name}: uus näit on väiksem kui eelmine '
                        f'({previous.value:g}). Kontrolli väärtust.'
                    )
                    continue

                reading = Reading.query.filter_by(
                    meter_id=meter.id,
                    period=period
                ).first()

                if reading:
                    reading.value = value
                    reading.day_value = None
                    reading.night_value = None
                    reading.note = note
                    reading.entered_by = current_user().id
                    reading.updated_at = datetime.utcnow()
                else:
                    reading = Reading(
                        meter_id=meter.id,
                        period=period,
                        value=value,
                        note=note,
                        entered_by=current_user().id
                    )
                    db.session.add(reading)

            saved += 1

        if saved:
            log_action(
                'readings_saved',
                f'{category.center.name} / {category.name} / '
                f'{period}: {saved} näitu'
            )
            db.session.commit()
            flash(f'Salvestatud näite: {saved}.', 'success')

        if errors:
            for error in errors:
                flash(error, 'error')

        return redirect(url_for(
            'category_readings',
            category_id=category.id,
            period=period
        ))

    rows = []

    for meter in meters:
        rows.append({
            'meter': meter,
            'reading': Reading.query.filter_by(
                meter_id=meter.id,
                period=period
            ).first(),
            'previous': previous_reading(meter.id, period)
        })

    return render_template(
        'category.html',
        category=category,
        rows=rows,
        period=period
    )
@app.route('/history')
@login_required
def history():
    period = request.args.get('period', date.today().strftime('%Y-%m'))
    if not valid_period(period): period = date.today().strftime('%Y-%m')
    center_id = request.args.get('center_id', type=int)
    query = Reading.query.join(Meter).join(Category).join(Center).filter(Reading.period == period)
    # Historical records remain visible through the month in which the relevant
    # object (center/category/meter) was removed, but never in later months.
    cutoff = period + '-01'
    query = query.filter((Center.deleted_at.is_(None)) | (db.func.strftime('%Y-%m', Center.deleted_at) >= period)) if db.engine.dialect.name == 'sqlite' else query.filter((Center.deleted_at.is_(None)) | (db.func.to_char(Center.deleted_at, 'YYYY-MM') >= period))
    query = query.filter((Category.deleted_at.is_(None)) | (db.func.strftime('%Y-%m', Category.deleted_at) >= period)) if db.engine.dialect.name == 'sqlite' else query.filter((Category.deleted_at.is_(None)) | (db.func.to_char(Category.deleted_at, 'YYYY-MM') >= period))
    query = query.filter((Meter.deleted_at.is_(None)) | (db.func.strftime('%Y-%m', Meter.deleted_at) >= period)) if db.engine.dialect.name == 'sqlite' else query.filter((Meter.deleted_at.is_(None)) | (db.func.to_char(Meter.deleted_at, 'YYYY-MM') >= period))
    if center_id: query = query.filter(Center.id == center_id)
    readings = query.order_by(Center.sort_order, Center.name, Category.sort_order, Meter.sort_order).all()
    centers = Center.query.order_by(Center.sort_order, Center.name).all()
    return render_template('history.html', readings=readings, centers=centers, period=period, center_id=center_id)

@app.route('/history/<int:reading_id>/delete', methods=['POST'])
@admin_required
def delete_reading(reading_id):
    reading = db.get_or_404(Reading, reading_id)

    period = reading.period
    center_id = reading.meter.category.center_id
    details = (
        f'{reading.meter.category.center.name} / '
        f'{reading.meter.category.name} / {reading.meter.name}; '
        f'period={period}; value={reading.value}; '
        f'day_value={reading.day_value}; night_value={reading.night_value}'
    )

    log_action('reading_deleted', details)
    db.session.delete(reading)
    db.session.commit()

    flash('Näidukirje kustutati jäädavalt.', 'success')
    return redirect(url_for(
        'history',
        period=period,
        center_id=center_id
    ))


@app.route('/history/<int:reading_id>/edit', methods=['POST'])
@admin_required
def edit_reading(reading_id):
    reading = db.get_or_404(Reading, reading_id)
    note = request.form.get('note', '').strip()
    previous = previous_reading(reading.meter_id, reading.period)

    try:
        if reading.meter.tariff_mode == 'dual':
            raw_day = request.form.get('day_value', '').strip().replace(' ', '').replace('\u00a0', '').replace(',', '.')
            raw_night = request.form.get('night_value', '').strip().replace(' ', '').replace('\u00a0', '').replace(',', '.')

            day_value = float(raw_day)
            night_value = float(raw_night)

            if (
                day_value < 0
                or night_value < 0
                or not math.isfinite(day_value)
                or not math.isfinite(night_value)
            ):
                raise ValueError()

            if previous and previous.day_value is not None:
                if day_value < previous.day_value:
                    flash(
                        'Päevanäit ei tohi olla väiksem eelmisest näidust.',
                        'error'
                    )
                    return redirect(url_for(
                        'history',
                        period=reading.period,
                        center_id=reading.meter.category.center_id
                    ))

            if previous and previous.night_value is not None:
                if night_value < previous.night_value:
                    flash(
                        'Öönäit ei tohi olla väiksem eelmisest näidust.',
                        'error'
                    )
                    return redirect(url_for(
                        'history',
                        period=reading.period,
                        center_id=reading.meter.category.center_id
                    ))

            old_value = reading.value
            reading.day_value = day_value
            reading.night_value = night_value
            reading.value = day_value + night_value

        else:
            raw_value = request.form.get('value', '').strip().replace(' ', '').replace('\u00a0', '').replace(',', '.')
            new_value = float(raw_value)

            if new_value < 0 or not math.isfinite(new_value):
                raise ValueError()

            if previous and new_value < previous.value:
                flash(
                    'Uus näit ei tohi olla väiksem eelmisest näidust.',
                    'error'
                )
                return redirect(url_for(
                    'history',
                    period=reading.period,
                    center_id=reading.meter.category.center_id
                ))

            old_value = reading.value
            reading.value = new_value
            reading.day_value = None
            reading.night_value = None

    except (ValueError, TypeError):
        flash('Sisesta korrektsed näidud.', 'error')
        return redirect(url_for(
            'history',
            period=reading.period,
            center_id=reading.meter.category.center_id
        ))

    old_note = reading.note or ''
    reading.note = note or None
    reading.updated_at = datetime.utcnow()

    log_action(
        'reading_edited',
        f'{reading.meter.category.center.name} / '
        f'{reading.meter.category.name} / {reading.meter.name}; '
        f'period={reading.period}; value={old_value} -> {reading.value}; '
        f'note_changed={old_note != (note or "")}'
    )
    db.session.commit()
    flash('Näidu muudatus salvestatud.', 'success')
    return redirect(url_for(
        'history',
        period=reading.period,
        center_id=reading.meter.category.center_id
    ))


@app.route('/export.xlsx')
@login_required
def export_xlsx():
    period = request.args.get('period', date.today().strftime('%Y-%m'))
    if not valid_period(period): period = date.today().strftime('%Y-%m')
    center_id = request.args.get('center_id', type=int)
    query = Reading.query.join(Meter).join(Category).join(Center).filter(Reading.period == period)
    if db.engine.dialect.name == 'sqlite':
        month_center = db.func.strftime('%Y-%m', Center.deleted_at)
        month_category = db.func.strftime('%Y-%m', Category.deleted_at)
        month_meter = db.func.strftime('%Y-%m', Meter.deleted_at)
    else:
        month_center = db.func.to_char(Center.deleted_at, 'YYYY-MM')
        month_category = db.func.to_char(Category.deleted_at, 'YYYY-MM')
        month_meter = db.func.to_char(Meter.deleted_at, 'YYYY-MM')
    query = query.filter((Center.deleted_at.is_(None)) | (month_center >= period))
    query = query.filter((Category.deleted_at.is_(None)) | (month_category >= period))
    query = query.filter((Meter.deleted_at.is_(None)) | (month_meter >= period))
    if center_id: query = query.filter(Center.id == center_id)
    readings = query.order_by(Center.sort_order, Center.name, Category.sort_order, Meter.sort_order).all()
    wb = Workbook(); ws = wb.active; ws.title = 'Näidud'
    ws.append([
        'Periood', 'Kaubanduskeskus', 'Kategooria', 'Mõõdik',
        'Mõõdiku kood', 'Ressurss', 'Ühik',
        'Eelmine näit', 'Praegune näit', 'Tarbitud',
        'Päeva eelmine näit', 'Päeva praegune näit', 'Päeva tarbimine',
        'Öö eelmine näit', 'Öö praegune näit', 'Öö tarbimine',
        'Märkus', 'Sisestaja', 'Sisestatud'
    ])
    def safe_excel_text(value):
        if not isinstance(value, str):
            return value
        probe = value.lstrip(' \\t\\r\\n')
        if probe.startswith(('=', '+', '-', '@')):
            return "'" + value
        return value

    for r in readings:
        prev = previous_reading(r.meter_id, period)
        if r.meter.tariff_mode == 'dual':
            day_consumption = (
                r.day_value - prev.day_value
                if prev and prev.day_value is not None and r.day_value is not None
                else None
            )
            night_consumption = (
                r.night_value - prev.night_value
                if prev and prev.night_value is not None and r.night_value is not None
                else None
            )
            ws.append([
                period, r.meter.category.center.name, r.meter.category.name,
                r.meter.name, r.meter.meter_code or '', r.meter.resource_type,
                r.meter.unit, None, r.value,
                (day_consumption + night_consumption)
                if day_consumption is not None and night_consumption is not None
                else None,
                prev.day_value if prev else None, r.day_value, day_consumption,
                prev.night_value if prev else None, r.night_value, night_consumption,
                r.note or '', r.user.username if r.user else '',
                r.entered_at.strftime('%Y-%m-%d %H:%M')
            ])
            for cell in ws[ws.max_row]:
                if isinstance(cell.value, str):
                    cell.value = safe_excel_text(cell.value)
        else:
            consumption = r.value - prev.value if prev else None
            ws.append([
                period, r.meter.category.center.name, r.meter.category.name,
                r.meter.name, r.meter.meter_code or '', r.meter.resource_type,
                r.meter.unit, prev.value if prev else None, r.value, consumption,
                None, None, None, None, None, None,
                r.note or '', r.user.username if r.user else '',
                r.entered_at.strftime('%Y-%m-%d %H:%M')
            ])
            for cell in ws[ws.max_row]:
                if isinstance(cell.value, str):
                    cell.value = safe_excel_text(cell.value)
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF'); cell.fill = PatternFill('solid', fgColor='174A6E'); cell.alignment = Alignment(wrap_text=True)
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        width = min(max(max(len(str(c.value or '')) for c in col) + 2, 12), 32)
        ws.column_dimensions[get_column_letter(col[0].column)].width = width
    output = io.BytesIO(); wb.save(output); output.seek(0)
    log_action('excel_export', f'period={period}; center={center_id or "all"}'); db.session.commit()
    filename = f'metervault_naidud_{period}' + (f'_center_{center_id}' if center_id else '') + '.xlsx'
    return send_file(output, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# Admin panel
@app.route('/admin')
@admin_required
def admin_home():
    centers = Center.query.filter(Center.deleted_at.is_(None)).order_by(Center.sort_order, Center.name).all()
    users = User.query.order_by(User.username).all()
    log_page = max(request.args.get('log_page', 1, type=int), 1)
    try:
        log_per_page = int(request.args.get('log_per_page', 10))
    except (TypeError, ValueError):
        log_per_page = 10
    if log_per_page not in {10, 30, 60, 100, 200}:
        log_per_page = 10
    log_pagination = AuditLog.query.order_by(AuditLog.created_at.desc()).paginate(
        page=log_page, per_page=log_per_page, error_out=False
    )
    return render_template('admin.html', centers=centers, users=users,
                           logs=log_pagination.items, log_pagination=log_pagination,
                           log_page=log_page, log_per_page=log_per_page)

@app.route('/admin/center', methods=['POST'])
@admin_required
def admin_add_center():
    name = request.form.get('name','').strip()
    if not name: flash('Sisesta keskuse nimi.', 'error')
    else:
        db.session.add(Center(name=name, sort_order=request.form.get('sort_order', 0, type=int)))
        log_action('center_created', name); db.session.commit(); flash('Kaubanduskeskus lisatud.', 'success')
    return redirect(url_for('admin_home'))

@app.route('/admin/center/<int:center_id>/edit', methods=['POST'])
@admin_required
def admin_edit_center(center_id):
    c = db.get_or_404(Center, center_id)
    if c.deleted_at is not None:
        flash('Kustutatud objekti ei saa muuta.', 'warning')
        return redirect(url_for('admin_home'))
    name = request.form.get('name','').strip()
    if name: c.name = name
    c.sort_order = request.form.get('sort_order', c.sort_order, type=int)
    log_action('center_updated', c.name); db.session.commit(); flash('Muudatused salvestatud.', 'success')
    return redirect(url_for('admin_home'))

@app.route('/admin/center/<int:center_id>/archive', methods=['POST'])
@admin_required
def admin_archive_center(center_id):
    c = db.get_or_404(Center, center_id)
    if c.deleted_at is not None:
        flash('Kustutatud objekti olekut ei saa muuta.', 'warning')
        return redirect(url_for('admin_home'))
    c.active = not c.active
    log_action('center_active_changed', f'{c.name}: active={c.active}')
    db.session.commit()
    flash('Keskuse olek muudetud.', 'success')
    return redirect(url_for('admin_home'))


@app.route('/admin/center/<int:center_id>/delete', methods=['POST'])
@admin_required
def admin_delete_center(center_id):
    c = db.get_or_404(Center, center_id)

    if c.deleted_at is not None:
        flash('Objekt on juba kustutatud.', 'info')
        return redirect(url_for('admin_home'))

    center_name = c.name
    deleted_time = datetime.utcnow()
    c.deleted_at = deleted_time
    c.active = False
    for category in c.categories:
        if category.deleted_at is None:
            category.deleted_at = deleted_time
        category.active = False
        for meter in category.meters:
            if meter.deleted_at is None:
                meter.deleted_at = deleted_time
            meter.active = False

    log_action(
        'center_deleted',
        f'{center_name}; historical readings preserved'
    )

    db.session.commit()
    flash(
        'Objekt eemaldati aktiivsest loendist. Kõik ajaloolised näidud säilitati.',
        'success'
    )
    return redirect(url_for('admin_home'))

@app.route('/admin/category', methods=['POST'])
@admin_required
def admin_add_category():
    name = request.form.get('name', '').strip()
    center_id = request.form.get('center_id', type=int)
    center = db.session.get(Center, center_id) if center_id else None

    if not name:
        flash('Sisesta kategooria nimi.', 'error')
    elif not center:
        flash('Vali kaubanduskeskus.', 'error')
    elif center.deleted_at is not None:
        flash('Kustutatud kaubanduskeskusele ei saa kategooriat lisada.', 'error')
    else:
        db.session.add(Category(
            center_id=center.id,
            name=name,
            comment=request.form.get('comment', '').strip()[:2000] or None,
            sort_order=request.form.get('sort_order', 0, type=int)
        ))
        log_action('category_created', f'{center.name} / {name}')
        db.session.commit()
        flash('Kategooria lisatud.', 'success')

    return redirect(url_for('admin_home'))

@app.route('/admin/category/<int:category_id>/edit', methods=['POST'])
@admin_required
def admin_edit_category(category_id):
    c = db.get_or_404(Category, category_id)
    name = request.form.get('name')
    if c.deleted_at is not None or c.center.deleted_at is not None:
        flash('Kustutatud kategooriat ei saa muuta.', 'warning')
    else:
        old_name = c.name
        old_comment = c.comment or ''
        if name is not None:
            name = name.strip()
            if not name:
                flash('Kategooria nimi ei saa olla tühi.', 'error')
                return redirect(url_for('admin_home'))
            c.name = name
            c.sort_order = request.form.get('sort_order', c.sort_order, type=int)
        if 'comment' in request.form:
            c.comment = request.form.get('comment', '').strip()[:2000] or None
        if old_name != c.name:
            log_action('category_updated', c.name)
        if old_comment != (c.comment or ''):
            log_action('category_comment_updated', f'{c.center.name} / {c.name}')
        db.session.commit()
        flash('Kategooria uuendatud.', 'success')
    return redirect(url_for('admin_home'))

@app.route('/admin/category/<int:category_id>/archive', methods=['POST'])
@admin_required
def admin_archive_category(category_id):
    c = db.get_or_404(Category, category_id)
    if c.deleted_at is not None:
        flash('Kustutatud kategooria olekut ei saa muuta.', 'warning')
    else:
        c.active = not c.active
        log_action('category_active_changed', f'{c.name}: active={c.active}')
        db.session.commit()
        flash('Kategooria olek muudetud.', 'success')
    return redirect(url_for('admin_home'))

@app.route('/admin/category/<int:category_id>/delete', methods=['POST'])
@admin_required
def admin_delete_category(category_id):
    c = db.get_or_404(Category, category_id)
    if c.deleted_at is None:
        deleted_time = datetime.utcnow()
        c.deleted_at = deleted_time
        c.active = False
        for meter in c.meters:
            if meter.deleted_at is None:
                meter.deleted_at = deleted_time
            meter.active = False
        log_action('category_deleted', f'{c.center.name} / {c.name}; historical readings preserved')
        db.session.commit()
        flash('Kategooria eemaldati aktiivsest arvestusest. Ajaloolised näidud säilitati.', 'success')
    return redirect(url_for('admin_home'))


@app.route('/admin/meter', methods=['POST'])
@admin_required
def admin_add_meter():
    category_id = request.form.get('category_id', type=int)
    category = db.session.get(Category, category_id) if category_id else None

    name = request.form.get('name', '').strip()
    resource = request.form.get('resource_type', 'Muu')
    unit = request.form.get('unit', '').strip()

    allowed = {'Elekter', 'Külm vesi', 'Soe vesi', 'Gaas', 'Muu'}

    tariff_mode = request.form.get('tariff_mode', 'single')
    if resource != 'Elekter':
        tariff_mode = 'single'

    if tariff_mode not in {'single', 'dual'}:
        tariff_mode = 'single'

    if not category or category.deleted_at is not None or category.center.deleted_at is not None or not name or resource not in allowed or not unit:
        flash('Kontrolli kategooriat, mõõdiku nime, ressurssi ja ühikut.', 'error')
    else:
        m = Meter(
            category_id=category.id,
            name=name,
            meter_code=request.form.get('meter_code', '').strip() or None,
            description=request.form.get('description', '').strip() or None,
            resource_type=resource,
            unit=unit,
            tariff_mode=tariff_mode,
            sort_order=request.form.get(
                'sort_order',
                max((meter.sort_order for meter in category.meters), default=-10) + 10,
                type=int
            )
        )

        db.session.add(m)
        log_action('meter_created', f'{category.center.name} / {category.name} / {name}')
        db.session.commit()
        flash('Mõõdik lisatud.', 'success')

    return redirect(url_for('admin_home'))


@app.route('/admin/category/<int:category_id>/meters/reorder', methods=['POST'])
@admin_required
def admin_reorder_meters(category_id):
    category = db.get_or_404(Category, category_id)
    payload = request.get_json(silent=True) or {}
    meter_ids = payload.get('meter_ids')

    if not isinstance(meter_ids, list) or not all(
        isinstance(meter_id, int) and not isinstance(meter_id, bool)
        for meter_id in meter_ids
    ):
        return jsonify({'success': False, 'message': 'Vigane mõõdikute järjestus.'}), 400

    # Deleted meters are intentionally absent from the admin list.
    # Accept the visible order and append any visible meters omitted by a stale page.
    visible_meters = [meter for meter in category.meters if meter.deleted_at is None]
    meters_by_id = {meter.id: meter for meter in visible_meters}

    if len(meter_ids) != len(set(meter_ids)):
        return jsonify({'success': False, 'message': 'Mõõdikute nimekirjas on korduvaid ID-sid.'}), 400

    if any(meter_id not in meters_by_id for meter_id in meter_ids):
        return jsonify({'success': False, 'message': 'Mõõdik ei kuulu sellesse kategooriasse või on kustutatud.'}), 400

    ordered_ids = list(meter_ids)
    ordered_ids.extend(
        meter.id for meter in visible_meters
        if meter.id not in set(ordered_ids)
    )

    for position, meter_id in enumerate(ordered_ids):
        meters_by_id[meter_id].sort_order = (position + 1) * 10

    log_action('meters_reordered', f'{category.center.name} / {category.name}')
    db.session.commit()

    return jsonify({
        'success': True,
        'saved_ids': ordered_ids
    })


@app.route('/admin/centers/reorder', methods=['POST'])
@admin_required
def admin_reorder_centers():
    payload = request.get_json(silent=True) or {}
    ids = payload.get('ids')
    visible = Center.query.filter(Center.deleted_at.is_(None)).all()
    mapping = {x.id: x for x in visible}
    if (not isinstance(ids, list)
        or not all(type(i) is int for i in ids)
        or len(ids) != len(set(ids))
        or any(i not in mapping for i in ids)):
        return jsonify({'success': False, 'message': 'Vigane objektide järjestus.'}), 400
    ids = ids + [x.id for x in visible if x.id not in ids]
    for n, ident in enumerate(ids, 1): mapping[ident].sort_order = n * 10
    log_action('centers_reordered', 'Kaubanduskeskuste järjestus muudetud')
    db.session.commit()
    return jsonify({'success': True})

@app.route('/admin/center/<int:center_id>/categories/reorder', methods=['POST'])
@admin_required
def admin_reorder_categories(center_id):
    center = db.get_or_404(Center, center_id)
    payload = request.get_json(silent=True) or {}
    ids = payload.get('ids')
    visible = [x for x in center.categories if x.deleted_at is None]
    mapping = {x.id: x for x in visible}
    if (not isinstance(ids, list)
        or not all(type(i) is int for i in ids)
        or len(ids) != len(set(ids))
        or any(i not in mapping for i in ids)):
        return jsonify({'success': False, 'message': 'Vigane kategooriate järjestus.'}), 400
    ids = ids + [x.id for x in visible if x.id not in ids]
    for n, ident in enumerate(ids, 1): mapping[ident].sort_order = n * 10
    log_action('categories_reordered', f'{center.name}')
    db.session.commit()
    return jsonify({'success': True})

@app.route('/admin/meter/<int:meter_id>/edit', methods=['POST'])
@admin_required
def admin_edit_meter(meter_id):
    m = db.get_or_404(Meter, meter_id)

    if m.deleted_at is not None:
        flash('Kustutatud mõõdikut ei saa muuta.', 'warning')
        return redirect(url_for('admin_home'))

    name = request.form.get('name', '').strip()
    unit = request.form.get('unit', '').strip()
    resource = request.form.get('resource_type', 'Muu')

    allowed = {'Elekter', 'Külm vesi', 'Soe vesi', 'Gaas', 'Muu'}

    tariff_mode = request.form.get('tariff_mode', 'single')
    if resource != 'Elekter':
        tariff_mode = 'single'

    if tariff_mode not in {'single', 'dual'}:
        tariff_mode = 'single'

    if name and unit and resource in allowed:
        m.name = name
        m.unit = unit
        m.resource_type = resource
        m.meter_code = request.form.get('meter_code', '').strip() or None
        m.description = request.form.get('description', '').strip() or None
        m.sort_order = request.form.get('sort_order', m.sort_order, type=int)
        m.tariff_mode = tariff_mode

        log_action('meter_updated', name)
        db.session.commit()
        flash('Mõõdik uuendatud.', 'success')
    else:
        flash('Kontrolli mõõdiku andmeid.', 'error')

    return redirect(url_for('admin_home'))




@app.route('/admin/meter/<int:meter_id>/archive', methods=['POST'])
@admin_required
def admin_archive_meter(meter_id):
    m = db.get_or_404(Meter, meter_id)

    if m.deleted_at is not None:
        flash('Kustutatud mõõdiku olekut ei saa muuta.', 'warning')
        return redirect(url_for('admin_home'))

    m.active = not m.active

    log_action(
        'meter_active_changed',
        f'{m.name}: active={m.active}'
    )

    db.session.commit()
    flash('Mõõdiku olek muudetud.', 'success')

    return redirect(url_for('admin_home'))


@app.route('/admin/meter/<int:meter_id>/delete', methods=['POST'])
@admin_required
def admin_delete_meter(meter_id):
    m = db.get_or_404(Meter, meter_id)

    if m.deleted_at is not None:
        flash('Mõõdik on juba kustutatud.', 'info')
        return redirect(url_for('admin_home'))

    meter_name = m.name
    category_name = m.category.name
    center_name = m.category.center.name

    m.deleted_at = datetime.utcnow()
    m.active = False

    log_action(
        'meter_deleted',
        f'{center_name} / {category_name} / {meter_name}; historical readings preserved'
    )

    db.session.commit()

    flash(
        'Mõõdik eemaldati aktiivsest arvestusest. Kõik ajaloolised näidud säilitati.',
        'success'
    )

    return redirect(url_for('admin_home'))

@app.route('/admin/user', methods=['POST'])
@admin_required
def admin_add_user():
    username=request.form.get('username','').strip(); password=request.form.get('password',''); role=request.form.get('role','user')
    if not username or len(password) < 14 or role not in {'user','admin'}:
        flash('Kasutajanimi on kohustuslik ja parool peab olema vähemalt 14 märki.', 'error')
    elif User.query.filter_by(username=username).first(): flash('See kasutajanimi on juba kasutusel.', 'error')
    else:
        u=User(username=username, role=role); u.set_password(password); db.session.add(u); log_action('user_created', f'{username} ({role})'); db.session.commit(); flash('Kasutaja lisatud.', 'success')
    return redirect(url_for('admin_home'))

@app.route('/admin/user/<int:user_id>/toggle', methods=['POST'])
@admin_required
def admin_toggle_user(user_id):
    u=db.get_or_404(User,user_id)
    if u.id == current_user().id: flash('Enda kontot ei saa siin deaktiveerida.', 'error')
    else:
        u.active=not u.active; log_action('user_active_changed', f'{u.username}: active={u.active}'); db.session.commit(); flash('Kasutaja olek muudetud.', 'success')
    return redirect(url_for('admin_home'))

@app.route('/admin/user/<int:user_id>/password', methods=['POST'])
@admin_required
def admin_reset_password(user_id):
    u=db.get_or_404(User,user_id); password=request.form.get('password','')
    if len(password) < 14: flash('Parool peab olema vähemalt 14 märki.', 'error')
    else: u.set_password(password); u.session_version += 1; log_action('password_reset', u.username); db.session.commit(); flash('Parool muudetud. Kasutaja peab uuesti sisse logima.', 'success')
    return redirect(url_for('admin_home'))

@app.errorhandler(403)
def forbidden(_): return render_template('error.html', code=403, message='Sul puudub sellele lehele ligipääs.'), 403
@app.errorhandler(404)
def not_found(_): return render_template('error.html', code=404, message='Lehte ei leitud.'), 404

with app.app_context():
    db.create_all()

    inspector = inspect(db.engine)

    user_columns = [column['name'] for column in inspector.get_columns('user')]
    if 'session_version' not in user_columns:
        db.session.execute(text(
            "ALTER TABLE \"user\" ADD COLUMN session_version INTEGER NOT NULL DEFAULT 1"
        ))

    center_columns = [column['name'] for column in inspector.get_columns('center')]

    if 'deleted_at' not in center_columns:
        db.session.execute(text(
            "ALTER TABLE center ADD COLUMN deleted_at TIMESTAMP"
        ))

    category_columns = [column['name'] for column in inspector.get_columns('category')]
    if 'deleted_at' not in category_columns:
        db.session.execute(text("ALTER TABLE category ADD COLUMN deleted_at TIMESTAMP"))

    category_columns = [column['name'] for column in inspector.get_columns('category')]
    if 'comment' not in category_columns:
        db.session.execute(text(
            "ALTER TABLE category ADD COLUMN comment TEXT"
        ))

    meter_columns = [column['name'] for column in inspector.get_columns('meter')]

    if 'deleted_at' not in meter_columns:
        db.session.execute(text(
            "ALTER TABLE meter ADD COLUMN deleted_at TIMESTAMP"
        ))
    if 'description' not in meter_columns:
        db.session.execute(text("ALTER TABLE meter ADD COLUMN description TEXT"))

    if 'tariff_mode' not in meter_columns:
        db.session.execute(text(
            "ALTER TABLE meter ADD COLUMN tariff_mode VARCHAR(10) NOT NULL DEFAULT 'single'"
        ))

    reading_columns = [column['name'] for column in inspector.get_columns('reading')]

    if 'day_value' not in reading_columns:
        db.session.execute(text(
            "ALTER TABLE reading ADD COLUMN day_value FLOAT"
        ))

    if 'night_value' not in reading_columns:
        db.session.execute(text(
            "ALTER TABLE reading ADD COLUMN night_value FLOAT"
        ))

    db.session.commit()
    # Optional first-run admin bootstrap. Never hardcode production credentials.
    bootstrap_user = os.environ.get('BOOTSTRAP_ADMIN_USER')
    bootstrap_password = os.environ.get('BOOTSTRAP_ADMIN_PASSWORD')
    if bootstrap_user and bootstrap_password:
        if len(bootstrap_user.strip()) > 80 or len(bootstrap_password) < 14:
            raise RuntimeError('Bootstrap admin username must be <=80 characters and password >=14 characters.')
        if not User.query.filter_by(role='admin').first():
            admin = User(username=bootstrap_user.strip(), role='admin')
            admin.set_password(bootstrap_password)
            db.session.add(admin)
            db.session.commit()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=False)
