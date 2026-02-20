from flask import Flask, render_template, request, redirect, url_for
from models import db, Expert, Committee, Membership, Meeting, Participation

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

if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(host="0.0.0.0", port=5000, debug=True)