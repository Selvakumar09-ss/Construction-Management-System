from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session, abort
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import os
import io
import csv
from datetime import datetime
from openpyxl import Workbook  # pip install openpyxl
from functools import wraps

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'static', 'uploads')
TEMPLATE_FOLDER = os.path.join(BASE_DIR, 'templates')
STATIC_FOLDER = os.path.join(BASE_DIR, 'static')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__, template_folder=TEMPLATE_FOLDER, static_folder=STATIC_FOLDER)
app.config['SECRET_KEY'] = 'change_this_secret_key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(BASE_DIR, 'app.db')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB upload limit

db = SQLAlchemy(app)

# ---------------------
# Models
# ---------------------
roles = ('Admin', 'Manager', 'Engineer', 'Worker', 'Client')

project_workers = db.Table('project_workers',
    db.Column('project_id', db.Integer, db.ForeignKey('project.id')),
    db.Column('worker_id', db.Integer, db.ForeignKey('user.id'))
)

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120))
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(200))
    role = db.Column(db.String(50), default='Worker')
    def set_password(self, pw):
        self.password_hash = generate_password_hash(pw)
    def check_password(self, pw):
        return check_password_hash(self.password_hash, pw)

class Project(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200))
    location = db.Column(db.String(200))
    start_date = db.Column(db.String(20))
    end_date = db.Column(db.String(20))
    budget_estimated = db.Column(db.Float, default=0.0)
    budget_actual = db.Column(db.Float, default=0.0)
    milestones = db.Column(db.Text)  # newline separated
    progress = db.Column(db.Integer, default=0)  # percentage
    client_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    workers = db.relationship('User', secondary=project_workers, backref='projects')

class Task(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200))
    description = db.Column(db.Text)
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'))
    assigned_to = db.Column(db.Integer, db.ForeignKey('user.id'))
    priority = db.Column(db.String(20), default='Normal')
    deadline = db.Column(db.String(20), nullable=True)
    status = db.Column(db.String(50), default='Pending')
    project = db.relationship('Project', backref='tasks', lazy=True)
    assigned_to_user = db.relationship('User', backref='assigned_tasks', lazy=True)


class Material(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200))
    quantity = db.Column(db.Float, default=0.0)
    unit = db.Column(db.String(50), default='pcs')
    threshold = db.Column(db.Float, default=0.0)
    used = db.Column(db.Float, default=0.0)

class Resource(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200))
    type = db.Column(db.String(100))  # machine/worker/equipment descriptor
    status = db.Column(db.String(100), default='available')

class Update(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'))
    author_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    text = db.Column(db.Text)
    timestamp = db.Column(db.String(50))
    images = db.Column(db.Text)  # comma-separated filenames

class Notification(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    message = db.Column(db.String(500))
    seen = db.Column(db.Boolean, default=False)
    created = db.Column(db.String(40))

# ---------------------
# Helpers
# ---------------------
def current_user():
    uid = session.get('user_id')
    if not uid:
        return None
    return User.query.get(uid)


def login_required(roles_allowed=None):
    def decorator(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            u = current_user()
            if not u:
                flash('Please login first', 'warning')
                return redirect(url_for('login'))
            if roles_allowed and u.role not in roles_allowed:
                flash('Access denied', 'danger')
                return redirect(url_for('index'))
            return fn(*args, **kwargs)
        return wrapper
    return decorator


def create_notification(message, user_id=None):
    """Create a notification. user_id=None means for all users (broadcast)."""
    n = Notification(
        user_id=user_id,
        message=message,
        seen=False,
        created=datetime.now().strftime('%Y-%m-%d %H:%M')
    )
    db.session.add(n)
    db.session.commit()


def user_can_access_project(user, project):
    """Check if user is allowed to view this project."""
    if not user:
        return False
    if user.role in ('Admin', 'Manager'):
        return True
    if user.role == 'Client':
        return project.client_id == user.id
    # Worker / Engineer: assigned via project_workers or tasks
    if project in user.projects:
        return True
    assigned_task = Task.query.filter_by(project_id=project.id, assigned_to=user.id).first()
    if assigned_task:
        return True
    return False


def get_user_projects(user):
    """Return only projects this user is allowed to see."""
    if not user:
        return []
    if user.role in ('Admin', 'Manager'):
        return Project.query.all()
    if user.role == 'Client':
        return Project.query.filter_by(client_id=user.id).all()
    # Worker / Engineer: projects they are assigned to
    projects = list(user.projects)
    task_pids = [t.project_id for t in Task.query.filter_by(assigned_to=user.id).all()]
    extra = Project.query.filter(Project.id.in_(task_pids)).all() if task_pids else []
    seen = {p.id for p in projects}
    for p in extra:
        if p.id not in seen:
            projects.append(p)
            seen.add(p.id)
    return projects


# ---------------------
# Routes: Auth & Index
# ---------------------
@app.route('/')
def index():
    u = current_user()
    if u:
        if u.role == 'Admin':
            return redirect(url_for('admin_dashboard'))
        elif u.role == 'Manager':
            return redirect(url_for('manager_dashboard'))
        elif u.role == 'Client':
            return redirect(url_for('client_view'))
        else:
            return redirect(url_for('projects'))
    return render_template('index.html')

@app.context_processor
def inject_user():
    u = current_user()
    unread = 0
    if u:
        unread = Notification.query.filter(
            Notification.user_id == u.id,
            Notification.seen == False
        ).count()
    return {'current_user': u, 'unread_count': unread}

@app.route('/register', methods=['GET','POST'])
@login_required(roles_allowed=['Admin'])
def register():
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        pw = request.form['password']
        role = request.form.get('role','Worker')
        if User.query.filter_by(email=email).first():
            flash('Email already exists', 'danger')
            return redirect(url_for('register'))
        u = User(name=name, email=email, role=role)
        u.set_password(pw)
        db.session.add(u)
        db.session.commit()
        flash('User registered successfully.', 'success')
        return redirect(url_for('admin_dashboard'))
    return render_template('register.html', roles=roles)

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method=='POST':
        email = request.form['email']
        pw = request.form['password']
        u = User.query.filter_by(email=email).first()
        if u and u.check_password(pw):
            session['user_id'] = u.id
            flash('Logged in successfully', 'success')
            return redirect(url_for('index'))
        else:
            flash('Invalid credentials', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('user_id', None)
    flash('Logged out', 'info')
    return redirect(url_for('index'))

@app.route('/change-password', methods=['GET', 'POST'])
def change_password():
    u = current_user()
    if not u:
        flash('Please login first', 'warning')
        return redirect(url_for('login'))
    if request.method == 'POST':
        current_pw = request.form.get('current_password', '')
        new_pw = request.form.get('new_password', '')
        confirm_pw = request.form.get('confirm_password', '')
        if not u.check_password(current_pw):
            flash('Current password is incorrect.', 'danger')
            return redirect(url_for('change_password'))
        if len(new_pw) < 6:
            flash('New password must be at least 6 characters.', 'warning')
            return redirect(url_for('change_password'))
        if new_pw != confirm_pw:
            flash('New passwords do not match.', 'danger')
            return redirect(url_for('change_password'))
        u.set_password(new_pw)
        db.session.commit()
        flash('Password changed successfully!', 'success')
        return redirect(url_for('index'))
    return render_template('change_password.html')

# ---------------------
# Admin Dashboard & User Management
# ---------------------
@app.route('/admin')
@login_required(roles_allowed=['Admin'])
def admin_dashboard():
    users = User.query.all()
    projects = Project.query.all()
    return render_template('dashboard_admin.html', users=users, projects=projects)

@app.route('/manager')
@login_required(roles_allowed=['Manager', 'Admin'])
def manager_dashboard():
    projects = Project.query.all()
    return render_template('project_list.html', projects=projects)

@app.route('/user/add', methods=['GET','POST'])
@login_required(roles_allowed=['Admin'])
def add_user():
    if request.method=='POST':
        name = request.form['name']
        email = request.form['email']
        role = request.form['role']
        pw = request.form['password']
        if User.query.filter_by(email=email).first():
            flash('Email already exists', 'danger')
            return redirect(url_for('add_user'))
        u = User(name=name, email=email, role=role)
        u.set_password(pw)
        db.session.add(u)
        db.session.commit()
        flash('User added', 'success')
        return redirect(url_for('admin_dashboard'))
    return render_template('user_form.html', roles=roles)

@app.route('/user/edit/<int:uid>', methods=['GET','POST'])
@login_required(roles_allowed=['Admin'])
def edit_user(uid):
    u = User.query.get_or_404(uid)
    if request.method=='POST':
        u.name = request.form['name']
        u.email = request.form['email']
        u.role = request.form['role']
        pw = request.form.get('password')
        if pw:
            u.set_password(pw)
        db.session.commit()
        flash('User updated', 'success')
        return redirect(url_for('admin_dashboard'))
    return render_template('user_form.html', user=u, roles=roles)

@app.route('/user/delete/<int:uid>')
@login_required(roles_allowed=['Admin'])
def delete_user(uid):
    u = User.query.get_or_404(uid)
    if u.id == current_user().id:
        flash('Cannot delete yourself', 'danger')
        return redirect(url_for('admin_dashboard'))
    db.session.delete(u)
    db.session.commit()
    flash('User deleted', 'success')
    return redirect(url_for('admin_dashboard'))

# ---------------------
# Worker Management
# ---------------------
@app.route('/workers')
@login_required()
def workers():
    list_workers = User.query.filter(User.role != 'Client').all()
    return render_template('worker_list.html', workers=list_workers)

# ---------------------
# Project CRUD
# ---------------------
@app.route('/projects')
@login_required()
def projects():
    u = current_user()
    projects = get_user_projects(u)
    return render_template('project_list.html', projects=projects)

@app.route('/project/add', methods=['GET','POST'])
@login_required(roles_allowed=['Admin','Manager'])
def add_project():
    if request.method=='POST':
        p = Project(
            name=request.form['name'],
            location=request.form['location'],
            start_date=request.form['start_date'],
            end_date=request.form['end_date'],
            budget_estimated=float(request.form.get('budget_estimated') or 0),
            milestones=request.form.get('milestones',''),
            client_id=int(request.form.get('client_id')) if request.form.get('client_id') else None
        )
        db.session.add(p)
        db.session.commit()
        # Assign workers if selected
        worker_ids = request.form.getlist('workers')
        for wid in worker_ids:
            w = User.query.get(int(wid))
            if w:
                p.workers.append(w)
        db.session.commit()
        if p.client_id:
            create_notification(f'You have been assigned to project: {p.name}', user_id=p.client_id)
        flash('Project created', 'success')
        return redirect(url_for('projects'))
    clients = User.query.filter_by(role='Client').all()
    workers = User.query.filter(User.role != 'Client').all()
    return render_template('project_form.html', clients=clients, workers=workers)

@app.route('/project/<int:pid>')
@login_required()
def project_detail(pid):
    project = Project.query.get_or_404(pid)
    u = current_user()
    if not user_can_access_project(u, project):
        flash('You do not have access to this project', 'danger')
        return redirect(url_for('projects'))
    tasks = Task.query.filter_by(project_id=pid).all()
    updates = Update.query.filter_by(project_id=pid).order_by(Update.id.desc()).all()
    return render_template('project_detail.html', project=project, tasks=tasks, updates=updates, User=User)

@app.route('/project/edit/<int:pid>', methods=['GET','POST'])
@login_required(roles_allowed=['Admin','Manager'])
def project_edit(pid):
    p = Project.query.get_or_404(pid)
    if request.method=='POST':
        p.name = request.form['name']
        p.location = request.form['location']
        p.start_date = request.form['start_date']
        p.end_date = request.form['end_date']
        p.budget_estimated = float(request.form.get('budget_estimated') or 0)
        p.budget_actual = float(request.form.get('budget_actual') or 0)
        p.milestones = request.form.get('milestones','')
        p.progress = int(request.form.get('progress') or 0)
        cid = request.form.get('client_id')
        p.client_id = int(cid) if cid else None
        db.session.commit()
        create_notification(f'Project updated: {p.name}')
        flash('Project updated', 'success')
        return redirect(url_for('project_detail', pid=pid))
    clients = User.query.filter_by(role='Client').all()
    return render_template('project_form.html', project=p, clients=clients)

@app.route('/project/delete/<int:pid>')
@login_required(roles_allowed=['Admin'])
def project_delete(pid):
    p = Project.query.get_or_404(pid)
    name = p.name
    db.session.delete(p)
    db.session.commit()
    create_notification(f'Project deleted: {name}')
    flash('Project deleted', 'success')
    return redirect(url_for('projects'))

# ---------------------
# Tasks
# ---------------------
@app.route('/tasks')
@login_required()
def task_list():
    u = current_user()
    if u.role == 'Client':
        projects = Project.query.filter_by(client_id=u.id).all()
        pids = [p.id for p in projects]
        tasks = Task.query.filter(Task.project_id.in_(pids)).all() if pids else []
    elif u.role in ('Worker', 'Engineer'):
        tasks = Task.query.filter_by(assigned_to=u.id).all()
    else:
        tasks = Task.query.all()
    return render_template('task_list.html', tasks=tasks)

@app.route('/task/add', methods=['GET','POST'])
@login_required(roles_allowed=['Admin','Manager'])
def add_task():
    if request.method=='POST':
        assigned = int(request.form.get('assigned_to')) if request.form.get('assigned_to') else None
        t = Task(
            title=request.form['title'],
            description=request.form.get('description',''),
            project_id=int(request.form['project_id']),
            assigned_to=assigned,
            priority=request.form.get('priority','Normal'),
            deadline=request.form.get('deadline'),
            status='Pending'
        )
        db.session.add(t)
        db.session.commit()
        if assigned:
            create_notification(f'You have been assigned task: {t.title}', user_id=assigned)
        flash('Task added', 'success')
        return redirect(url_for('task_list'))
    projects = Project.query.all()
    workers = User.query.filter(User.role.in_(['Worker', 'Engineer', 'Manager'])).all()
    return render_template('task_form.html', projects=projects, workers=workers)

@app.route('/task/edit/<int:tid>', methods=['GET','POST'])
@login_required()
def edit_task(tid):
    t = Task.query.get_or_404(tid)
    u = current_user()
    # Restrict access
    if u.role not in ('Admin', 'Manager') and t.assigned_to != u.id:
        flash('Access denied', 'danger')
        return redirect(url_for('task_list'))
    if request.method=='POST':
        old_status = t.status
        t.title = request.form['title']
        t.description = request.form.get('description','')
        t.priority = request.form.get('priority','Normal')
        t.deadline = request.form.get('deadline')
        t.status = request.form.get('status','Pending')
        db.session.commit()
        if old_status != t.status:
            if t.assigned_to:
                create_notification(f'Your task "{t.title}" is now: {t.status}', user_id=t.assigned_to)
        flash('Task updated', 'success')
        return redirect(url_for('task_list'))
    projects = Project.query.all()
    workers = User.query.filter(User.role.in_(['Worker', 'Engineer', 'Manager'])).all()
    return render_template('task_form.html', task=t, projects=projects, workers=workers)

@app.route('/task/delete/<int:tid>')
@login_required(roles_allowed=['Admin'])
def delete_task(tid):
    t = Task.query.get_or_404(tid)
    title = t.title
    db.session.delete(t)
    db.session.commit()
    create_notification(f'Task deleted: {title}')
    flash('Task deleted', 'success')
    return redirect(url_for('task_list'))

# ---------------------
# Material Inventory
# ---------------------
@app.route('/materials')
@login_required()
def materials():
    mats = Material.query.all()
    low = [m for m in mats if m.quantity - m.used <= m.threshold]
    return render_template('material_list.html', materials=mats, low=low)

@app.route('/material/add', methods=['GET','POST'])
@login_required(roles_allowed=['Admin','Manager'])
def add_material():
    if request.method=='POST':
        m = Material(
            name=request.form['name'],
            quantity=float(request.form.get('quantity') or 0),
            unit=request.form.get('unit','pcs'),
            threshold=float(request.form.get('threshold') or 0)
        )
        db.session.add(m)
        db.session.commit()
        flash('Material added', 'success')
        return redirect(url_for('materials'))
    return render_template('material_form.html')

@app.route('/material/update/<int:mid>', methods=['GET','POST'])
@login_required(roles_allowed=['Admin','Manager'])
def update_material(mid):
    m = Material.query.get_or_404(mid)
    if request.method=='POST':
        m.quantity = float(request.form.get('quantity') or m.quantity)
        m.used = float(request.form.get('used') or m.used)
        db.session.commit()
        remaining = m.quantity - m.used
        if remaining <= m.threshold:
            create_notification(f'Low stock alert: {m.name} (remaining: {remaining} {m.unit})')
        flash('Material updated', 'success')
        return redirect(url_for('materials'))
    return render_template('material_form.html', material=m)

# ---------------------
# Resources
# ---------------------
@app.route('/resources')
@login_required()
def resources():
    res = Resource.query.all()
    return render_template('resource_list.html', resources=res)

@app.route('/resource/add', methods=['GET','POST'])
@login_required(roles_allowed=['Admin','Manager'])
def add_resource():
    if request.method=='POST':
        r = Resource(name=request.form['name'], type=request.form.get('type','machine'))
        db.session.add(r)
        db.session.commit()
        flash('Resource added', 'success')
        return redirect(url_for('resources'))
    return render_template('resource_form.html')

# ---------------------
# Project Updates with Media
# ---------------------
def save_images(files):
    saved = []
    for f in files:
        if not f or not f.filename:
            continue
        filename = secure_filename(f.filename)
        ts = datetime.now().strftime('%Y%m%d%H%M%S%f')
        filename = f"{ts}_{filename}"
        path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        f.save(path)
        saved.append(filename)
    return saved

@app.route('/project/<int:pid>/update', methods=['GET','POST'])
@login_required(roles_allowed=['Admin','Manager'])
def project_update(pid):
    p = Project.query.get_or_404(pid)
    if request.method=='POST':
        text = request.form.get('text','')
        files = request.files.getlist('images')
        imgs = save_images(files)
        up = Update(
            project_id=pid,
            author_id=current_user().id,
            text=text,
            timestamp=datetime.now().strftime('%Y-%m-%d %H:%M'),
            images=",".join(imgs)
        )
        db.session.add(up)
        db.session.commit()
        # Notify only the assigned client, not broadcast to all
        if p.client_id:
            create_notification(f'New update on your project "{p.name}"', user_id=p.client_id)
        # Notify assigned workers only
        for worker in p.workers:
            if worker.id != current_user().id:
                create_notification(f'New update posted for project: {p.name}', user_id=worker.id)
        flash('Update posted', 'success')
        return redirect(url_for('project_detail', pid=pid))
    return render_template('upload_update.html', project=p)

# ---------------------
# Client view
# ---------------------
@app.route('/client')
@login_required(roles_allowed=['Client'])
def client_view():
    u = current_user()
    projects = Project.query.filter_by(client_id=u.id).all()
    return render_template('client_view.html', projects=projects)

# ---------------------
# Reports: CSV / Excel basic
# ---------------------
@app.route('/reports')
@login_required(roles_allowed=['Admin','Manager'])
def reports():
    return render_template('reports.html')

@app.route('/reports/download/<string:kind>')
@login_required(roles_allowed=['Admin','Manager'])
def download_report(kind):
    if kind == 'projects_xlsx':
        wb = Workbook()
        ws = wb.active
        ws.title = "Projects"
        ws.append(['ID','Name','Location','Start','End','Estimated','Actual','Progress'])
        for p in Project.query.all():
            ws.append([p.id,p.name,p.location,p.start_date,p.end_date,p.budget_estimated,p.budget_actual,p.progress])
        stream = io.BytesIO()
        wb.save(stream)
        stream.seek(0)
        return send_file(stream, as_attachment=True, download_name='projects.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    elif kind == 'materials_csv':
        si = io.StringIO()
        cw = csv.writer(si)
        cw.writerow(['ID','Name','Quantity','Used','Unit','Threshold'])
        for m in Material.query.all():
            cw.writerow([m.id,m.name,m.quantity,m.used,m.unit,m.threshold])
        mem = io.BytesIO()
        mem.write(si.getvalue().encode('utf-8'))
        mem.seek(0)
        return send_file(mem, as_attachment=True, download_name='materials.csv', mimetype='text/csv')
    flash('Unknown report', 'warning')
    return redirect(url_for('reports'))

# ---------------------
# Notifications
# ---------------------
@app.route('/notifications')
@login_required()
def notifications():
    u = current_user()
    notes = Notification.query.filter(
        Notification.user_id == u.id
    ).order_by(Notification.id.desc()).all()
    # Mark all as seen
    for n in notes:
        if not n.seen:
            n.seen = True
    db.session.commit()
    return render_template('notifications.html', notes=notes)

# ---------------------
# Image Download Route
# ---------------------
@app.route('/download/image/<filename>')
@login_required()
def download_image(filename):
    safe_name = secure_filename(filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_name)
    if not os.path.exists(file_path):
        abort(404)
    return send_file(file_path, as_attachment=True, download_name=safe_name)

# ---------------------
# Init DB helper
# ---------------------
with app.app_context():
    db.create_all()
    print("database created or already exists!")
    # Auto-create/update admin user
    admin = User.query.filter_by(email='admin123@gmail.com').first()
    if not admin:
        admin = User(name='Admin', email='admin123@gmail.com', role='Admin')
        admin.set_password('admin123')
        db.session.add(admin)
        db.session.commit()
        print("Admin user created: admin123@gmail.com / admin123")
    else:
        admin.set_password('admin123')
        db.session.commit()
        print("Admin password reset to: admin123")

if __name__ == '__main__':
    app.run(debug=False)

