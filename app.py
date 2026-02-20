from flask import Flask, render_template, request, redirect, url_for
from models import db, Expert, Committee, Membership, Meeting, Participation
import csv
from flask import Response

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///committee.db'
app.config['SECRET_KEY'] = 'yoursecretkey'
db.init_app(app)

@app.route('/')
def dashboard():
    experts = Expert.query.all()
    committees = Committee.query.all()
    meetings = Meeting.query.all()
    participations = Participation.query.all()
    return render_template(
        'dashboard.html',
        experts=experts,
        committees=committees,
        meetings=meetings,
        participations=participations
    )

@app.route('/add_expert', methods=['GET', 'POST'])
def add_expert():
    committees = Committee.query.all()
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        committee_id = request.form.get('committee_id')

        expert = Expert(name=name, email=email)
        db.session.add(expert)
        db.session.commit()

        # If committee selected, create membership
        if committee_id:
            membership = Membership(expert_id=expert.id, committee_id=committee_id)
            db.session.add(membership)
            db.session.commit()

        return redirect(url_for('dashboard'))
    return render_template('experts.html', committees=committees)

    @app.route('/add_meeting', methods=['GET', 'POST'])
def add_meeting():
    committees = Committee.query.all()
    if request.method == 'POST':
        committee_id = request.form['committee_id']
        date = request.form['date']
        agenda = request.form['agenda']

        meeting = Meeting(
            committee_id=committee_id,
            date=date,
            agenda=agenda
        )
        db.session.add(meeting)
        db.session.commit()

        return redirect(url_for('dashboard'))
    return render_template('meetings.html', committees=committees)

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

    return redirect(url_for('dashboard'))

    @app.route('/update_participation/<int:participation_id>', methods=['POST'])
def update_participation(participation_id):
    participation = Participation.query.get_or_404(participation_id)

    participation.attendance = 'attendance' in request.form
    participation.report_submitted = 'report_submitted' in request.form
    participation.reminder_sent = 'reminder_sent' in request.form

    db.session.commit()
    return redirect(url_for('dashboard'))

    @app.route('/delete_participation/<int:participation_id>', methods=['POST'])
def delete_participation(participation_id):
    participation = Participation.query.get_or_404(participation_id)
    db.session.delete(participation)
    db.session.commit()
    return redirect(url_for('dashboard'))


@app.route('/export_participation/<int:committee_id>')
def export_participation(committee_id):
    committee = Committee.query.get_or_404(committee_id)
    meetings = Meeting.query.filter_by(committee_id=committee_id).all()

    # Create CSV response
    def generate():
        data = []
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

if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(host="0.0.0.0", port=5000, debug=True)