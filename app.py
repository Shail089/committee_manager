from flask import Flask, render_template, request, redirect, url_for, Response, flash, send_file
from models import db, Expert, Committee, Membership, Meeting, Participation, NationalMirrorCommittee
from apscheduler.schedulers.background import BackgroundScheduler
import csv
from datetime import date, datetime
from openpyxl import Workbook
import io
from emails import (
    announcement_email,
    reminder_email_all,
    reminder_email_individual,
    completion_email,
    request_update_email
)

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///committee.db'
app.config['SECRET_KEY'] = 'yoursecretkey'
db.init_app(app)


@app.context_processor
def inject_current_year():
    return {'current_year': date.today().year}

@app.route('/')
def homepage():
    # High-level summary view
    nmcs = NationalMirrorCommittee.query.all()
    scs = Committee.query.filter(Committee.parent_id == None).all()
    wgs = Committee.query.filter(Committee.parent_id != None).all()
    experts = Expert.query.all()

    today = date.today()

    # ✅ Only future meetings
    meetings = Meeting.query.filter(Meeting.date >= today).order_by(Meeting.date.asc()).all()
    participations = Participation.query.all()

    return render_template(
        'index.html',
        nmcs=nmcs,
        scs=scs,
        wgs=wgs,
        experts=experts,
        meetings=meetings,
        participations=participations,
        current_year=today.year   # ✅ pass year into template
    )

@app.route('/dashboard')
def dashboard():
    nmcs = NationalMirrorCommittee.query.all()
    today = date.today()
    nmc_meetings = {}
    upcoming_count = 0
    past_count = 0

    for nmc in nmcs:
        meetings_set = set()
        for sc in nmc.subcommittees:
            meetings_set.update(sc.meetings)
            for wg in sc.children:
                meetings_set.update(wg.meetings)

        upcoming = [m for m in meetings_set if m.date >= today]
        past = [m for m in meetings_set if m.date < today]

        upcoming_count += len(upcoming)
        past_count += len(past)

        nmc_meetings[nmc.id] = {
            "upcoming": sorted(upcoming, key=lambda m: m.date),
            "past": sorted(past, key=lambda m: m.date, reverse=True)
        }

    participations = Participation.query.all()
    return render_template(
        'dashboard.html',
        nmcs=nmcs,
        nmc_meetings=nmc_meetings,
        participations=participations,
        upcoming_count=upcoming_count,
        past_count=past_count
    )
    
@app.route('/directory')
def directory():
    # Management page
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
        code = request.form['code']          # e.g. "WG 9"
        title = request.form['title']
        sc_id = request.form['sc_id']

        # Fetch the SC object
        sc = Committee.query.get(sc_id)
        if not sc:
            flash("Invalid SC selected", "danger")
            return redirect(url_for('directory'))

        # Auto‑prefix SC code to WG code
        merged_code = f"{sc.code}/{code}"

        wg = Committee(
            code=merged_code,
            title=title,
            parent_id=sc.id,     # ✅ link WG to SC
            nmc_id=sc.nmc_id     # ✅ inherit NMC from SC
        )
        db.session.add(wg)
        db.session.commit()

        flash('WG added successfully!', 'success')
        return redirect(url_for('directory'))

    return render_template('add_wg.html', nmcs=nmcs, scs=scs)

@app.route('/add_membership', methods=['GET', 'POST'])
def add_membership():
    nmcs = NationalMirrorCommittee.query.all()
    experts = Expert.query.all()
    scs = Committee.query.filter(Committee.parent_id == None).all()
    wgs = Committee.query.filter(Committee.parent_id != None).all()

    if request.method == 'POST':
        expert_id = request.form['expert_id']
        committee_id = request.form['committee_id']

        membership = Membership(expert_id=expert_id, committee_id=committee_id)
        db.session.add(membership)
        db.session.commit()

        flash('Membership added successfully!', 'success')
        return redirect(url_for('directory'))

    return render_template('add_membership.html', nmcs=nmcs, experts=experts, scs=scs, wgs=wgs)

# Add Expert
@app.route('/add_expert', methods=['POST'])
def add_expert():
    name = request.form['name']
    email = request.form['email']
    mobile = request.form.get('mobile')
    organisation = request.form.get('organisation')

    new_expert = Expert(
        name=name,
        email=email,
        mobile=mobile,
        organisation=organisation
    )
    db.session.add(new_expert)
    db.session.commit()

    flash('Expert added successfully!', 'success')
    # ✅ Redirect back to Update Directory page
    return redirect(url_for('directory'))

@app.route('/get_scs/<int:nmc_id>')
def get_scs(nmc_id):
    scs = Committee.query.filter_by(nmc_id=nmc_id, parent_id=None).all()
    return jsonify([{'id': sc.id, 'code': sc.code, 'title': sc.title} for sc in scs])

@app.route('/get_wgs/<int:sc_id>')
def get_wgs(sc_id):
    sc = Committee.query.get(sc_id)
    wgs = sc.children if sc else []
    return jsonify([{'id': wg.id, 'code': wg.code, 'title': wg.title} for wg in wgs])

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
from datetime import date, datetime
from flask import request, redirect, url_for, flash, render_template
from flask_mail import Mail, Message

from flask_mail import Mail, Message

app.config['MAIL_SERVER'] = 'smtp.mgovcloud.in'   # or smtp.zoho.in / smtp.zoho.eu depending on your domain
app.config['MAIL_PORT'] = 587                 # TLS port
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USE_SSL'] = False
app.config['MAIL_USERNAME'] = 'sailendra@bis.gov.in'   # full Zoho Gov email
app.config['MAIL_PASSWORD'] = '72dNMNWT9fvw'        # the 16-char app password
app.config['MAIL_DEFAULT_SENDER'] = 'sailendra@bis.gov.in'

mail = Mail(app)

# Add Meeting Route
@app.route('/add_meeting', methods=['GET', 'POST'])
def add_meeting():
    if request.method == 'POST':
        committee_id = request.form['committee_id']
        date_str = request.form['date']
        agenda = request.form['agenda']

        try:
            meeting_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            flash("Invalid date format. Please use YYYY-MM-DD.", "danger")
            return redirect(url_for('add_meeting'))

        meeting = Meeting(committee_id=committee_id, date=meeting_date, agenda=agenda)
        db.session.add(meeting)
        db.session.commit()

        # Create participations
        memberships = Membership.query.filter_by(committee_id=committee_id).all()
        for m in memberships:
            participation = Participation(meeting_id=meeting.id, expert_id=m.expert_id)
            db.session.add(participation)
        db.session.commit()

        # Collect recipient emails
        participations = Participation.query.filter_by(meeting_id=meeting.id).all()
        recipient_emails = [p.expert.email for p in participations]

        # Send announcement email
        msg = announcement_email(meeting, recipient_emails)
        mail.send(msg)

        flash('Meeting scheduled, participation prepared, and announcement email sent!', 'success')
        return redirect(url_for('add_meeting'))

    # Prepare upcoming/past meetings for template
    scs = Committee.query.filter_by(parent_id=None).all()
    nmcs = NationalMirrorCommittee.query.all()
    today = date.today()
    nmc_meetings = {}
    for nmc in nmcs:
        meetings_set = set()
        for sc in nmc.subcommittees:
            meetings_set.update(sc.meetings)
            for wg in sc.children:
                meetings_set.update(wg.meetings)

        upcoming = [m for m in meetings_set if m.date >= today]
        past = [m for m in meetings_set if m.date < today]

        nmc_meetings[nmc.id] = {
            "upcoming": sorted(upcoming, key=lambda m: m.date),
            "past": sorted(past, key=lambda m: m.date, reverse=True)
        }

    return render_template('meetings.html', scs=scs, nmcs=nmcs, nmc_meetings=nmc_meetings)


# Send Reminder to all participants
@app.route('/send_reminder/<int:meeting_id>', methods=['POST'])
def send_reminder(meeting_id):
    meeting = Meeting.query.get_or_404(meeting_id)
    participations = Participation.query.filter_by(meeting_id=meeting.id).all()
    recipient_emails = [p.expert.email for p in participations]

    msg = reminder_email_all(meeting, recipient_emails)
    mail.send(msg)

    for p in participations:
        p.reminder_sent = True
    db.session.commit()

    flash('Reminder sent to all members in one email!', 'info')
    return redirect(url_for('dashboard'))


# Send individual reminder
@app.route('/send_individual_reminder/<int:participation_id>', methods=['POST'])
def send_individual_reminder(participation_id):
    p = Participation.query.get_or_404(participation_id)
    meeting = p.meeting

    msg = reminder_email_individual(meeting, p.expert)
    mail.send(msg)

    p.reminder_sent = True
    db.session.commit()

    flash(f'Reminder sent to {p.expert.name}!', 'info')
    return redirect(url_for('dashboard'))


# Send completion email (after meeting)
@app.route('/send_completion/<int:meeting_id>', methods=['POST'])
def send_completion(meeting_id):
    meeting = Meeting.query.get_or_404(meeting_id)
    participations = Participation.query.filter_by(meeting_id=meeting.id).all()
    recipient_emails = [p.expert.email for p in participations]

    msg = completion_email(meeting, recipient_emails)
    mail.send(msg)

    flash('Completion email sent to all experts!', 'success')
    return redirect(url_for('dashboard'))


# Send request update email (to individual expert after completion)
@app.route('/send_request_update/<int:participation_id>', methods=['POST'])
def send_request_update_route(participation_id):
    p = Participation.query.get_or_404(participation_id)
    meeting = p.meeting

    msg = request_update_email(meeting, p.expert)
    mail.send(msg)

    flash(f'Participation status request sent to {p.expert.name}!', 'info')
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
    # Redirect back to dashboard with past tab active
    return redirect(url_for('dashboard', tab='past'))

# Delete Participation
@app.route('/delete_participation/<int:participation_id>', methods=['POST'])
def delete_participation(participation_id):
    participation = Participation.query.get_or_404(participation_id)
    db.session.delete(participation)
    db.session.commit()
    flash('Participation deleted!', 'danger')
    return redirect(url_for('dashboard'))

# Export for a single committee (NMC)
@app.route('/export_participation/<int:committee_id>')
def export_participation(committee_id):
    committee = NationalMirrorCommittee.query.get_or_404(committee_id)

    wb = Workbook()
    ws = wb.active
    ws.title = "Participation Report"
    # Add export date at the very top
    ws.append([f"Export Date: {date.today().strftime('%Y-%m-%d')}"])
    ws.append([])  # blank row for spacing
    # Header row
    ws.append([
        "NMC Code", "NMC Title",
        "SC Code", "SC Title",
        "WG Code", "WG Title",
        "Meeting Date", "Agenda",
        "Expert", "Organisation", "Email", "Phone",
        "Participation", "Report Submitted", "Reminder Sent"
    ])

    for sc in committee.subcommittees:
        for wg in sc.children:
            for meeting in wg.meetings:
                
                if meeting.date < date.today():
                    participations = Participation.query.filter_by(meeting_id=meeting.id).all()
                    for p in participations:
                        ws.append([
                            committee.code,
                            committee.title,
                            sc.code,
                            sc.title,
                            wg.code,
                            wg.title,
                            meeting.date.strftime("%Y-%m-%d"),
                            meeting.agenda or "",
                            p.expert.name,
                            p.expert.organisation or "",
                            p.expert.email or "",
                            p.expert.mobile or "",
                            "Yes" if p.attendance else "No",
                            "Yes" if p.report_submitted else "No",
                            "Yes" if p.reminder_sent else "No"
                        ])

    # Save to memory
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    # ✅ Add export date to filename
    export_date = date.today().strftime("%Y-%m-%d")
    filename = f"{committee.code}_participation_{export_date}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Export for ALL committees
@app.route('/export_all_participation')
def export_all_participation():
    nmcs = NationalMirrorCommittee.query.all()

    wb = Workbook()
    ws = wb.active
    ws.title = "All Participation"
    # Add export date at the very top
    ws.append([f"Export Date: {date.today().strftime('%Y-%m-%d')}"])
    ws.append([])  # blank row for spacing
    # Header row
    ws.append([
        "NMC Code", "NMC Title",
        "SC Code", "SC Title",
        "WG Code", "WG Title",
        "Meeting Date", "Agenda",
        "Expert", "Organisation", "Email", "Phone",
        "Participation", "Report Submitted", "Reminder Sent"
    ])

    for nmc in nmcs:
        for sc in nmc.subcommittees:
            for wg in sc.children:
                for meeting in wg.meetings:
                    if meeting.date < date.today():
                        participations = Participation.query.filter_by(meeting_id=meeting.id).all()
                        for p in participations:
                            ws.append([
                                nmc.code,
                                nmc.title,
                                sc.code,
                                sc.title,
                                wg.code,
                                wg.title,
                                meeting.date.strftime("%Y-%m-%d"),
                                meeting.agenda or "",
                                p.expert.name,
                                p.expert.organisation or "",
                                p.expert.email or "",
                                p.expert.mobile or "",
                                "Yes" if p.attendance else "No",
                                "Yes" if p.report_submitted else "No",
                                "Yes" if p.reminder_sent else "No"
                            ])

    # Save to memory
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
     # ✅ Add export date to filename
    export_date = date.today().strftime("%Y-%m-%d")
    filename = f"all_committees_participation_{export_date}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
# export experts with their NMC and SC/WG memberships
@app.route('/export_experts')
def export_experts():
    experts = Expert.query.all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Experts"

    # Header row
    ws.append(["Name", "Email", "Phone", "Organisation", "NMC", "Nominations (SC/WG)"])

    for expert in experts:
        # Group memberships by NMC
        nmc_map = {}
        for m in expert.memberships:
            committee = m.committee
            # Walk up hierarchy until top-level NMC
            while committee.parent_id is not None:
                committee = committee.parent
            nmc_code = committee.nmc.code  # ✅ top-level NMC code
            if nmc_code not in nmc_map:
                nmc_map[nmc_code] = []
            nmc_map[nmc_code].append(f"{m.committee.code} - {m.committee.title}")

        # Write one row per NMC
        if nmc_map:
            for nmc_code, nominations in nmc_map.items():
                ws.append([
                    expert.name,
                    expert.email,
                    expert.mobile or "",
                    expert.organisation or "",
                    nmc_code,
                    "; ".join(nominations)
                ])
        else:
            # Expert with no memberships
            ws.append([
                expert.name,
                expert.email,
                expert.mobile or "",
                expert.organisation or "",
                "None",
                "None"
            ])

    # Save to memory
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="experts_summary.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Seed Data
from datetime import date, timedelta

@app.route('/seed')
def seed_data():
    if not NationalMirrorCommittee.query.first():
        # Create NMCs
        nmc1 = NationalMirrorCommittee(code="LITD 19", title="E-Learning")
        nmc2 = NationalMirrorCommittee(
            code="LITD 24",
            title="Magnetic Components, Ferrite Materials, Piezoelectric and Frequency Control Devices"
        )
        db.session.add_all([nmc1, nmc2])
        db.session.commit()

        # Create SCs
        sc1 = Committee(
            code="ISO/IEC JTC 1/SC 36",
            title="Information technology for learning, education and training",
            nmc_id=nmc1.id,
            parent_id=None
        )
        sc2 = Committee(
            code="IEC/TC 49",
            title="Piezoelectric, dielectric and electrostatic devices and associated materials for frequency control, selection and detection",
            nmc_id=nmc2.id,
            parent_id=None
        )
        sc3 = Committee(
            code="IEC/TC 51",
            title="Magnetic components, ferrite and magnetic powder materials",
            nmc_id=nmc2.id,
            parent_id=None
        )
        db.session.add_all([sc1, sc2, sc3])
        db.session.commit()

        # Create WGs
        wg1 = Committee(code=f"{sc1.code}/WG 3", title="Learner information", parent_id=sc1.id, nmc_id=nmc1.id)
        wg2 = Committee(code=f"{sc1.code}/WG 7", title="Culture, language and individual needs", parent_id=sc1.id, nmc_id=nmc1.id)
        wg3 = Committee(code=f"{sc3.code}/WG 9", title="Inductive components", parent_id=sc3.id, nmc_id=nmc2.id)
        wg4 = Committee(code=f"{sc3.code}/WG 10", title="Magnetic materials and components for EMC applications", parent_id=sc3.id, nmc_id=nmc2.id)
        wg5 = Committee(code=f"{sc2.code}/WG 7", title="Piezoelectric, dielectric and electrostatic oscillators", parent_id=sc2.id, nmc_id=nmc2.id)
        db.session.add_all([wg1, wg2, wg3, wg4, wg5])
        db.session.commit()

        # Experts
        e1 = Expert(name="Dr. Alice", email="alice@example.com", mobile="1234567890", organisation="IIT Delhi")
        e2 = Expert(name="Dr. Bob", email="bob@example.com", mobile="9876543210", organisation="IIT Madras")
        e3 = Expert(name="Dr. Charlie", email="charlie@example.com", mobile="5555555555", organisation="BIS")
        db.session.add_all([e1, e2, e3])
        db.session.commit()

        # Memberships
        m1 = Membership(expert_id=e1.id, committee_id=sc1.id)
        m2 = Membership(expert_id=e2.id, committee_id=wg1.id)
        m3 = Membership(expert_id=e3.id, committee_id=sc2.id)
        m4 = Membership(expert_id=e3.id, committee_id=wg3.id)
        db.session.add_all([m1, m2, m3, m4])
        db.session.commit()

        # Meetings (some past, some upcoming)
        today = date.today()
        past1 = Meeting(committee_id=sc1.id, date=today - timedelta(days=30), agenda="Review of e-learning standards")
        past2 = Meeting(committee_id=sc2.id, date=today - timedelta(days=10), agenda="Discussion on piezoelectric devices")
        upcoming1 = Meeting(committee_id=wg1.id, date=today + timedelta(days=7), agenda="Learner information schema update")
        upcoming2 = Meeting(committee_id=wg3.id, date=today + timedelta(days=15), agenda="Inductive components testing protocols")
        upcoming3 = Meeting(committee_id=wg5.id, date=today + timedelta(days=45), agenda="Oscillator material improvements")

        db.session.add_all([past1, past2, upcoming1, upcoming2, upcoming3])
        db.session.commit()

    flash("Database seeded successfully with committees, experts, memberships, and meetings!", "success")
    return redirect(url_for('homepage'))

def send_completion_emails():
    today = date.today()
    past_meetings = Meeting.query.filter(
        Meeting.date < today,
        Meeting.completion_sent == False
    ).all()

    for meeting in past_meetings:
        participations = Participation.query.filter_by(meeting_id=meeting.id).all()
        recipient_emails = [p.expert.email for p in participations]

        msg = completion_email(meeting, recipient_emails)
        mail.send(msg)

        meeting.completion_sent = True
        db.session.commit()

        print(f"Completion email sent for meeting {meeting.id}")

# --- Scheduler setup ---
scheduler = BackgroundScheduler()
scheduler.add_job(send_completion_emails, 'interval', days=1)  # run once per day
scheduler.start()

if __name__ == "__main__":
    with app.app_context():
        #db.drop_all()
        db.create_all()
    app.run(host="0.0.0.0", port=5000, debug=True)
