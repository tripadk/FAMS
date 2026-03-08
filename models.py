from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

class Faculty(db.Model):
    faculty_id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(100), unique=True, nullable=False)
    password = db.Column(db.String(100), nullable=False)
    department = db.Column(db.String(100), nullable=False)
    designation = db.Column(db.String(100), nullable=False)
    profile_photo = db.Column(db.String(200))
    achievements = db.relationship('Achievements', backref='faculty', lazy=True)

class Achievements(db.Model):
    achievement_id = db.Column(db.Integer, primary_key=True)
    faculty_id = db.Column(db.Integer, db.ForeignKey('faculty.faculty_id'), nullable=False)
    title = db.Column(db.String(200), nullable=False)
    type = db.Column(db.String(100), nullable=False)
    description = db.Column(db.Text, nullable=False)
    date = db.Column(db.Date, nullable=False)
    proof_file = db.Column(db.String(200))
    certificate_file = db.Column(db.String(200))
    status = db.Column(db.String(20), default='Pending')

class Admin(db.Model):
    admin_id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(100), unique=True, nullable=False)
    password = db.Column(db.String(100), nullable=False)