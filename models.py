from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

class NationalMirrorCommittee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(50), unique=True, nullable=False)
    title = db.Column(db.String(200), nullable=False)
    subcommittees = db.relationship('Committee', backref='nmc', lazy=True)

class Committee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(50), unique=True, nullable=False)
    title = db.Column(db.String(200), nullable=False)
    parent_id = db.Column(db.Integer, db.ForeignKey('committee.id'), nullable=True)
    parent = db.relationship('Committee', remote_side=[id])
    nmc_id = db.Column(db.Integer, db.ForeignKey('national_mirror_committee.id'), nullable=False)

class Expert(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100))
    email = db.Column(db.String(120), unique=True)
    mobile = db.Column(db.String(20))          # new field
    organisation = db.Column(db.String(200))   # new field

class Membership(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    expert_id = db.Column(db.Integer, db.ForeignKey('expert.id'))
    committee_id = db.Column(db.Integer, db.ForeignKey('committee.id'))
    expert = db.relationship('Expert', backref='memberships')
    committee = db.relationship('Committee', backref='memberships')

class Meeting(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    committee_id = db.Column(db.Integer, db.ForeignKey('committee.id'), nullable=False)
    date = db.Column(db.Date, nullable=False)
    agenda = db.Column(db.Text, nullable=True)
    committee = db.relationship('Committee', backref='meetings')

class Participation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    meeting_id = db.Column(db.Integer, db.ForeignKey('meeting.id'))
    expert_id = db.Column(db.Integer, db.ForeignKey('expert.id'))
    attendance = db.Column(db.Boolean, default=False)
    report_submitted = db.Column(db.Boolean, default=False)
    reminder_sent = db.Column(db.Boolean, default=False)
    meeting = db.relationship('Meeting')
    expert = db.relationship('Expert')
