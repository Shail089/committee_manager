from flask import Flask, render_template, request, redirect, url_for, Response, flash
from models import db, Expert, Committee, Membership, Meeting, Participation, NationalMirrorCommittee
import csv
from datetime import date, datetime

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///committee.db'
app.config['SECRET_KEY'] = 'yoursecretkey'
db.init_app(app)

# Dashboard - NMC → SC → WG hierarchy
@app.route('/')
def dashboard():
    nmcs = NationalMirrorCommittee.query.all()
    experts = Expert.query.all()
    meetings = Meeting.query.all()
    participations = Participation.query.all()
    return render_template(
        'dashboard.html',
        nmcs=nmcs,
        experts=experts,
        meetings=meetings,
        participations=participations
    )
    
@app.route('/directory')
def directory():
    nmcs = NationalMirrorCommittee.query.all()
    scs = Committee.query.filter(Committee.parent_id == None).all()
    wgs = Committee.query.filter(Committee.parent_id != None).all()
    experts = Expert.query.all()
    memberships = Membership.query.all()

    return render_template(
        'directory.html',
        nmcs=nmcs,
        scs=scs,
        wgs=wgs,
        experts=experts,
        memberships=memberships
    )
    
@app.route('/add_nmc', methods=['GET', 'POST'])
def add_nmc():
    if request.method == 'POST':
        code = request.form['code']
        title = request.form['title']
        nmc = NationalMirrorCommittee(code=code, title=title)
        db.session.add(nmc)
        db.session.commit()
        flash('NMC added successfully!', 'success')
        return redirect(url_for('directory'))
    return render_template('add_nmc.html')

@app.route('/add_sc', methods=['GET', 'POST'])
def add_sc():
    nmcs = NationalMirrorCommittee.query.all()
    if request.method == 'POST':
        code = request.form['code']
        title = request.form['title']
        nmc_id = request.form['nmc_id']
        sc = Committee(code=code, title=title, nmc_id=nmc_id)
        db.session.add(sc)
        db.session.commit()
        flash('SC added successfully!', 'success')
        return redirect(url_for('directory'))
    return render_template('add_sc.html', nmcs=nmcs)

@app.route('/add_wg', methods=['GET', 'POST'])
def add_wg():
    nmcs = NationalMirrorCommittee.query.all()
    scs = Committee.query.filter(Committee.parent_id == None).all()  # only SCs
    if request.method == 'POST':
        code = request.form['code']
        title = request.form['title']
        nmc_id = request.form['nmc_id']
        sc_id = request.form['sc_id']
        wg = Committee(code=code, title=title, parent_id=sc_id, nmc_id=nmc_id)
        db.session.add(wg)
        db.session.commit()
        flash('WG added successfully!', 'success')
        return redirect(url_for('directory'))
    return render_template('add_wg.html', nmcs=nmcs, scs=scs)

@app.route('/add_membership', methods=['GET', 'POST'])
def add_membership():
    nmcs = NationalMirrorCommittee.query.all()
    scs = Committee.query.filter(Committee.parent_id == None).all()
    wgs = Committee.query.filter(Committee.parent_id != None).all()
    experts = Expert.query.all()

    if request.method == 'POST':
        expert_id = request.form['expert_id']
        sc_id = request.form.get('sc_id')
        wg_id = request.form.get('wg_id')

        if sc_id and not wg_id:
            membership = Membership(expert_id=expert_id, committee_id=sc_id)
            db.session.add(membership)
        elif wg_id and not sc_id:
            membership = Membership(expert_id=expert_id, committee_id=wg_id)
            db.session.add(membership)
        elif sc_id and wg_id:
            m1 = Membership(expert_id=expert_id, committee_id=sc_id)
            m2 = Membership(expert_id=expert_id, committee_id=wg_id)
            db.session.add_all([m1, m2])

        db.session.commit()
        flash('Membership added successfully!', 'success')
        return redirect(url_for('directory'))

    return render_template('add_membership.html', nmcs=nmcs, scs=scs, wgs=wgs, experts=experts)

# Add Expert
@app.route('/add_expert', methods=['GET', 'POST'])
def add_expert():
    committees = Committee.query.all()
    experts = Expert.query.all()
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        mobile = request.form.get('mobile')
        organisation = request.form.get('organisation')
        committee_id = request.form.get('committee_id')

        expert = Expert(name=name, email=email, mobile=mobile, organisation=organisation)
        db.session.add(expert)
        db.session.commit()

        if committee_id:
            membership = Membership(expert_id=expert.id, committee_id=committee_id)
            db.session.add(membership)
            db.session.commit()

        flash('Expert added successfully!', 'success')
        return redirect(url_for('add_expert'))
    return render_template('experts.html', committees=committees, experts=experts)

# Edit Expert
@app.route('/edit_expert/<int:expert_id>', methods=['GET', 'POST'])
def edit_expert(expert_id):
    expert = Expert.query.get_or_404(expert_id)
    committees = Committee.query.all()
    if request.method == 'POST':
        expert.name = request.form['name']
        expert.email = request.form['email']
        expert.mobile = request.form.get('mobile')
        expert.organisation = request.form.get('organisation')
        committee_id = request.form.get('committee_id')
        if committee_id:
            Membership.query.filter_by(expert_id=expert.id).delete()
            new_membership = Membership(expert_id=expert.id, committee_id=committee_id)
            db.session.add(new_membership)
        db.session.commit()
        flash('Expert updated successfully!', 'success')
        return redirect(url_for('add_expert'))
    return render_template('edit_expert.html', expert=expert, committees=committees)

# Delete Expert
@app.route('/delete_expert/<int:expert_id>', methods=['POST'])
def delete_expert(expert_id):
    expert = Expert.query.get_or_404(expert_id)
    Membership.query.filter_by(expert_id=expert.id).delete()
    db.session.delete(expert)
    db.session.commit()
    flash('Expert deleted successfully!', 'danger')
    return redirect(url_for('add_expert'))

# Add Meeting
@app.route('/add_meeting', methods=['GET', 'POST'])
def add_meeting():
    committees = Committee.query.all()
    if request.method == 'POST':
        committee_id = request.form['committee_id']
        date_str = request.form['date']
        agenda = request.form['agenda']

        try:
            meeting_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            flash("Invalid date format. Please use YYYY-MM-DD.", "danger")
            return redirect(url_for('add_meeting'))

        # Create meeting
        meeting = Meeting(committee_id=committee_id, date=meeting_date, agenda=agenda)
        db.session.add(meeting)
        db.session.commit()

        # Auto-populate participation for all experts in this committee
        memberships = Membership.query.filter_by(committee_id=committee_id).all()
        for m in memberships:
            participation = Participation(meeting_id=meeting.id, expert_id=m.expert_id)
            db.session.add(participation)
        db.session.commit()

        flash('Meeting scheduled and participation table prepared!', 'success')
        return redirect(url_for('dashboard'))

    return render_template('meetings.html', committees=committees)

@app.route('/send_reminder/<int:meeting_id>', methods=['POST'])
def send_reminder(meeting_id):
    meeting = Meeting.query.get_or_404(meeting_id)
    participations = Participation.query.filter_by(meeting_id=meeting.id).all()
    for p in participations:
        send_email(
            subject=f"Reminder: {meeting.committee.title} Meeting on {meeting.date}",
            recipients=[p.expert.email],
            body=f"Dear {p.expert.name},\n\nThis is a reminder for the meeting on {meeting.date}.\nAgenda: {meeting.agenda}\n\nRegards,\nCommittee Manager"
        )
        p.reminder_sent = True
    db.session.commit()
    flash('Reminders sent to all members!', 'info')
    return redirect(url_for('dashboard'))

@app.route('/send_individual_reminder/<int:participation_id>', methods=['POST'])
def send_individual_reminder(participation_id):
    p = Participation.query.get_or_404(participation_id)
    meeting = p.meeting
    send_email(
        subject=f"Reminder: {meeting.committee.title} Meeting on {meeting.date}",
        recipients=[p.expert.email],
        body=f"Dear {p.expert.name},\n\nThis is a reminder for the meeting on {meeting.date}.\nAgenda: {meeting.agenda}\n\nRegards,\nCommittee Manager"
    )
    p.reminder_sent = True
    db.session.commit()
    flash(f'Reminder sent to {p.expert.name}!', 'info')
    return redirect(url_for('dashboard'))

# Add Participation
@app.route('/add_participation/<int:meeting_id>', methods=['POST'])
def add_participation(meeting_id):
    expert_id = request.form['expert_id']
    attendance = 'attendance' in request.form
    report_submitted = 'report_submitted' in request.form
    reminder_sent = 'reminder_sent' in request.form
    participation = Participation(
        meeting_id=meeting_id,
        expert_id=expert_id,
        attendance=attendance,
        report_submitted=report_submitted,
        reminder_sent=reminder_sent
    )
    db.session.add(participation)
    db.session.commit()
    flash('Participation added!', 'success')
    return redirect(url_for('dashboard'))

# Update Participation
@app.route('/update_participation/<int:participation_id>', methods=['POST'])
def update_participation(participation_id):
    participation = Participation.query.get_or_404(participation_id)
    participation.attendance = 'attendance' in request.form
    participation.report_submitted = 'report_submitted' in request.form
    participation.reminder_sent = 'reminder_sent' in request.form
    db.session.commit()
    flash('Participation updated!', 'info')
    return redirect(url_for('dashboard'))

# Delete Participation
@app.route('/delete_participation/<int:participation_id>', methods=['POST'])
def delete_participation(participation_id):
    participation = Participation.query.get_or_404(participation_id)
    db.session.delete(participation)
    db.session.commit()
    flash('Participation deleted!', 'danger')
    return redirect(url_for('dashboard'))

# Export Participation
@app.route('/export_participation/<int:committee_id>')
def export_participation(committee_id):
    committee = Committee.query.get_or_404(committee_id)
    meetings = Meeting.query.filter_by(committee_id=committee_id).all()
    def generate():
        header = ['Meeting Date', 'Expert', 'Attendance', 'Report Submitted', 'Reminder Sent']
        yield ','.join(header) + '\n'
        for meeting in meetings:
            participations = Participation.query.filter_by(meeting_id=meeting.id).all()
            for p in participations:
                row = [
                    str(meeting.date),
                    p.expert.name,
                    'Yes' if p.attendance else 'No',
                    'Yes' if p.report_submitted else 'No',
                    'Yes' if p.reminder_sent else 'No'
                ]
                yield ','.join(row) + '\n'
    return Response(generate(), mimetype='text/csv',
                    headers={"Content-Disposition": f"attachment;filename={committee.name}_participation.csv"})

# Seed Data
@app.route('/seed')
def seed_data():
    if not NationalMirrorCommittee.query.first():
        # Create NMCs
        nmc1 = NationalMirrorCommittee(code="LITD 19", title="Quantum Technologies")
        nmc2 = NationalMirrorCommittee(code="LITD 24", title="Electronics")
        db.session.add_all([nmc1, nmc2])
        db.session.commit()

        # Create SCs mapped to NMCs
        sc1 = Committee(code="SC19", title="SC Quantum Technologies", nmc=nmc1)
        sc2 = Committee(code="SC24", title="SC Electronics", nmc=nmc2)

        # Create WGs mapped to SCs and NMCs
        wg1 = Committee(code="WG19A", title="WG Cryptography", parent=sc1, nmc=nmc1)
        wg2 = Committee(code="WG19B", title="WG Quantum Communication", parent=sc1, nmc=nmc1)
        wg3 = Committee(code="WG24A", title="WG Microchips", parent=sc2, nmc=nmc2)

        db.session.add_all([sc1, sc2, wg1, wg2, wg3])
        db.session.commit()

        # Experts
        e1 = Expert(name="Dr. Alice", email="alice@example.com", mobile="1234567890", organisation="IIT Delhi")
        e2 = Expert(name="Dr. Bob", email="bob@example.com", mobile="9876543210", organisation="IIT Madras")
        e3 = Expert(name="Dr. Charlie", email="charlie@example.com", mobile="5555555555", organisation="BIS")
        db.session.add_all([e1, e2, e3])
        db.session.commit()

        # Memberships
        m1 = Membership(expert_id=e1.id, committee_id=sc1.id)   # SC only
        m2 = Membership(expert_id=e2.id, committee_id=wg1.id)   # WG only
        m3 = Membership(expert_id=e3.id, committee_id=sc2.id)   # SC only
        m4 = Membership(expert_id=e3.id, committee_id=wg3.id)   # WG also
        db.session.add_all([m1, m2, m3, m4])
        db.session.commit()

        return "Seed data added!"
    else:
        return "Database already has data."

if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(host="0.0.0.0", port=5000, debug=True)
