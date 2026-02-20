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
    memberships = Membership.query.all()
    return render_template(
        'dashboard.html',
        experts=experts,
        committees=committees,
        meetings=meetings,
        memberships=memberships
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

if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(host="0.0.0.0", port=5000, debug=True)