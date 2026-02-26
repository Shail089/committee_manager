from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

class NationalMirrorCommittee(db.Model):
    __tablename__ = "national_mirror_committee"
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(50), unique=True, nullable=False)
    title = db.Column(db.String(200), nullable=False)

    # All committees (SCs + WGs) under this NMC
    subcommittees = db.relationship("Committee", backref="nmc", lazy=True)


class Committee(db.Model):
    __tablename__ = "committee"
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(100), unique=True, nullable=False)
    title = db.Column(db.String(200), nullable=False)

    parent_id = db.Column(db.Integer, db.ForeignKey("committee.id"), nullable=True)
    parent = db.relationship("Committee", remote_side=[id], backref="children")

    nmc_id = db.Column(db.Integer, db.ForeignKey("national_mirror_committee.id"), nullable=False)

    memberships = db.relationship("Membership", back_populates="committee", lazy=True)


class Expert(db.Model):
    __tablename__ = "expert"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100))
    email = db.Column(db.String(120), unique=True)
    mobile = db.Column(db.String(20))
    organisation = db.Column(db.String(200))

    memberships = db.relationship("Membership", back_populates="expert", lazy=True)

    def get_nmc_map(self):
        """Return a dict {NMC_code: [SC/WG nominations]}"""
        nmc_map = {}
        for m in self.memberships:
            committee = m.committee
            # Walk up hierarchy until top-level NMC
            while committee.parent_id is not None:
                committee = committee.parent
            nmc_code = committee.nmc.code
            if nmc_code not in nmc_map:
                nmc_map[nmc_code] = []
            nmc_map[nmc_code].append(f"{m.committee.code} - {m.committee.title}")
        return nmc_map


class Membership(db.Model):
    __tablename__ = "membership"
    id = db.Column(db.Integer, primary_key=True)
    expert_id = db.Column(db.Integer, db.ForeignKey("expert.id"))
    committee_id = db.Column(db.Integer, db.ForeignKey("committee.id"))

    expert = db.relationship("Expert", back_populates="memberships")
    committee = db.relationship("Committee", back_populates="memberships")


class Meeting(db.Model):
    __tablename__ = "meeting"
    id = db.Column(db.Integer, primary_key=True)
    committee_id = db.Column(db.Integer, db.ForeignKey("committee.id"), nullable=False)
    date = db.Column(db.Date, nullable=False)
    agenda = db.Column(db.Text, nullable=True)

    committee = db.relationship("Committee", backref="meetings")

    # NEW FIELD
    completion_sent = db.Column(db.Boolean, default=False)


class Participation(db.Model):
    __tablename__ = "participation"
    id = db.Column(db.Integer, primary_key=True)
    meeting_id = db.Column(db.Integer, db.ForeignKey("meeting.id"))
    expert_id = db.Column(db.Integer, db.ForeignKey("expert.id"))
    attendance = db.Column(db.Boolean, default=False)
    report_submitted = db.Column(db.Boolean, default=False)
    reminder_sent = db.Column(db.Boolean, default=False)

    meeting = db.relationship("Meeting")
    expert = db.relationship("Expert")