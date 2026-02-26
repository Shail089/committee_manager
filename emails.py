from flask_mail import Message
from datetime import timedelta

# 1. Announcement Email (when meeting is added)
def announcement_email(meeting, recipients):
    return Message(
        subject=f"Announcement: {meeting.committee.title} Meeting Scheduled",
        recipients=recipients,
        html=f"""
            <p><strong>Dear Experts,</strong></p>
            <p>The meeting of <strong>{meeting.committee.code} - {meeting.committee.title}</strong> 
            has been scheduled on <strong>{meeting.date}</strong>. 
            Please ensure participation and prepare reports for 
            <strong>{meeting.committee.nmc.code}-{meeting.committee.nmc.title}</strong>.</p>
            <p><strong>Best regards</strong><br>
            <strong>Sailendra Kumar Verma</strong><br>
            <strong>Member Secretary, {meeting.committee.nmc.code}</strong></p>
        """
    )


# 2. Reminder Email (to all experts)
def reminder_email_all(meeting, recipients):
    return Message(
        subject=f"Reminder: {meeting.committee.title} Meeting on {meeting.date}",
        recipients=recipients,
        html=f"""
            <p><strong>Dear Experts,</strong></p>
            <p>This is a reminder that the meeting of 
            <strong>{meeting.committee.code} - {meeting.committee.title}</strong> will be held on 
            <strong>{meeting.date}</strong>. Please ensure your participation and prepare your report for 
            <strong>{meeting.committee.nmc.code}-{meeting.committee.nmc.title}</strong>.</p>
            <p><strong>Best regards</strong><br>
            <strong>Sailendra Kumar Verma</strong><br>
            <strong>Member Secretary, {meeting.committee.nmc.code}</strong></p>
        """
    )


# 3. Reminder Email (to individual expert)
def reminder_email_individual(meeting, expert):
    return Message(
        subject=f"Reminder: {meeting.committee.title} Meeting on {meeting.date}",
        recipients=[expert.email],
        html=f"""
            <p><strong>Dear {expert.name},</strong></p>
            <p>This is a reminder that the meeting of 
            <strong>{meeting.committee.code} - {meeting.committee.title}</strong> will be held on 
            <strong>{meeting.date}</strong>. As a registered expert, your participation is required and you are expected to present your report to 
            <strong>{meeting.committee.nmc.code}-{meeting.committee.nmc.title}</strong>.</p>
            <p><strong>Best regards</strong><br>
            <strong>Sailendra Kumar Verma</strong><br>
            <strong>Member Secretary, {meeting.committee.nmc.code}</strong></p>
        """
    )


# 4. Completion Email (after meeting is done, to all experts)
def completion_email(meeting, recipients):
    return Message(
        subject=f"Follow-up: {meeting.committee.title} Meeting Completed",
        recipients=recipients,
        html=f"""
            <p><strong>Dear Experts,</strong></p>
            <p>The meeting of <strong>{meeting.committee.code} - {meeting.committee.title}</strong> held on 
            <strong>{meeting.date}</strong> has been completed. You are requested to submit your participation status and reports to 
            <strong>{meeting.committee.nmc.code}-{meeting.committee.nmc.title}</strong>.</p>
            <p><strong>Best regards</strong><br>
            <strong>Sailendra Kumar Verma</strong><br>
            <strong>Member Secretary, {meeting.committee.nmc.code}</strong></p>
        """
    )


# 5. Request Update Email (to individual expert after completion)
def request_update_email(meeting, expert):
    deadline = meeting.date + timedelta(days=10)
    return Message(
        subject=f"Action Required: Submit Report for {meeting.committee.title}",
        recipients=[expert.email],
        html=f"""
            <p><strong>Dear {expert.name},</strong></p>
            <p>Following the meeting of <strong>{meeting.committee.code} - {meeting.committee.title}</strong> held on 
            <strong>{meeting.date}</strong>, you are requested to submit your participation status and report to 
            <strong>{meeting.committee.nmc.code}-{meeting.committee.nmc.title}</strong> by <strong>{deadline.strftime('%d %B %Y')}</strong>.</p>
            <p><strong>Best regards</strong><br>
            <strong>Sailendra Kumar Verma</strong><br>
            <strong>Member Secretary, {meeting.committee.nmc.code}</strong></p>
        """
    )