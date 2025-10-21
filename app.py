import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from datetime import datetime

st.set_page_config(
    page_title="Student Class Completion System",
    page_icon="✉️",
    layout="wide"
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 0.375rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 0.375rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 0.375rem;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def validate_app_password(password):
    """Validate Gmail App Password format"""
    if not password:
        return False, "Password cannot be empty"
    
    clean_password = password.replace(" ", "")
    if len(clean_password) != 16:
        return False, f"App Password should be 16 characters (got {len(clean_password)})"
    
    return True, "Valid App Password format"

def send_email(sender_email, sender_password, recipient_email, student_name, subject, body):
    """Send email to student about class completion"""
    is_valid, validation_msg = validate_app_password(sender_password)
    if not is_valid:
        return False, f"Invalid App Password: {validation_msg}"
    
    clean_password = sender_password.replace(" ", "")
    
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    
    msg.attach(MIMEText(body, 'plain'))
    
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.set_debuglevel(0)
    server.ehlo()
    server.starttls()
    server.ehlo()
    
    server.login(sender_email, clean_password)
    
    text = msg.as_string()
    server.sendmail(sender_email, recipient_email, text)
    server.quit()
    
    return True, "Email sent successfully"

def process_excel_file(uploaded_file):
    """Process uploaded Excel file and identify students with zero percentage"""
    df = pd.read_excel(uploaded_file)
    
    required_columns = ['Name', 'Email', 'Percentage']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        st.error(f"❌ Missing required columns: {', '.join(missing_columns)}")
        return None
    
    zero_percentage_students = df[df['Percentage'] == 0]
    
    return zero_percentage_students

def main():
    st.title("📚 Student Class Completion System")
    
    st.subheader("📧 Email Configuration")
    col1, col2 = st.columns(2)
    
    with col1:
        sender_email = st.text_input("Sender Email", value="robarthook009@gmail.com")
    
    with col2:
        sender_password = st.text_input("Sender Password", type="password")
    
    st.subheader("📁 Upload Excel File")
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if st.button("📤 Send Emails to Students with 0%", type="primary"):
        if not uploaded_file:
            st.error("Please upload an Excel file first")
        elif not sender_email or not sender_password:
            st.error("Please enter both email and password")
        else:
            zero_percentage_students = process_excel_file(uploaded_file)
            
            if zero_percentage_students is not None and len(zero_percentage_students) > 0:
                email_subject = "Important: Complete Your Classes - Action Required"
                email_template = """Dear {student_name},

We noticed that you haven't completed your classes yet (0% completion). 

Please log in to your account and complete the remaining classes as soon as possible. Your progress is important for your academic success.

If you have any questions or need assistance, please don't hesitate to contact us.

Best regards,
Innovative Skills LTD"""
                
                progress_bar = st.progress(0)
                success_count = 0
                error_count = 0
                
                for i, (_, student) in enumerate(zero_percentage_students.iterrows()):
                    progress_bar.progress((i + 1) / len(zero_percentage_students))
                    
                    personalized_body = email_template.format(student_name=student['Name'])
                    
                    success, message = send_email(
                        sender_email,
                        sender_password,
                        student['Email'],
                        student['Name'],
                        email_subject,
                        personalized_body
                    )
                    
                    if success:
                        success_count += 1
                        st.success(f"✅ Email sent to {student['Name']} ({student['Email']})")
                    else:
                        error_count += 1
                        st.error(f"❌ Failed to send email to {student['Name']}: {message}")
                
                if success_count > 0:
                    st.success(f"🎉 Successfully sent {success_count} emails!")
                if error_count > 0:
                    st.error(f"⚠️ Failed to send {error_count} emails.")
            else:
                st.info("No students with 0% completion found.")

if __name__ == "__main__":
    main()
