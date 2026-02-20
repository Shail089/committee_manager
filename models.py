from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

class Committee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), unique=True)
    parent_id = db.Column(db.Integer, db.ForeignKey('committee.id'), nullable=True)
    parent = db.relationship('Committee', remote_side=[id])

class Expert(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100))
    email = db.Column(db.String(120), unique=True)

class Membership(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    expert_id = db.Column(db.Integer, db.ForeignKey('expert.id'))
    committee_id = db.Column(db.Integer, db.ForeignKey('committee.id'))
    expert = db.relationship('Expert', backref='memberships')
    committee = db.relationship('Committee', backref='memberships')

class Meeting(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    committee_id = db.Column(db.Integer, db.ForeignKey('committee.id'))
    date = db.Column(db.Date)
    committee = db.relationship('Committee')

class Participation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    meeting_id = db.Column(db.Integer, db.ForeignKey('meeting.id'))
    expert_id = db.Column(db.Integer, db.ForeignKey('expert.id'))
    attendance = db.Column(db.Boolean, default=False)
    report_submitted = db.Column(db.Boolean, default=False)
    reminder_sent = db.Column(db.Boolean, default=False)
    meeting = db.relationship('Meeting')
    expert = db.relationship('Expert')