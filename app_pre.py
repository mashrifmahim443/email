import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from datetime import datetime

def validate_app_password(password):
    if not password:
        return False, "Password cannot be empty"
    
    clean_password = password.replace(" ", "")
    if len(clean_password) != 16:
        return False, f"App Password should be 16 characters (got {len(clean_password)})"
    
    return True, "Valid App Password format"

def send_email(sender_email, sender_password, recipient_email, student_name, subject, body):
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

def process_excel_file(file_path):
    df = pd.read_excel(file_path)
    
    required_columns = ['Name', 'Email', 'Percentage']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        print(f"Missing required columns: {', '.join(missing_columns)}")
        return None
    
    zero_percentage_students = df[df['Percentage'] == 0]
    
    return zero_percentage_students

def main():
    sender_email = "robarthook009@gmail.com"
    sender_password = input("Enter your Gmail App Password: ")
    
    file_path = input("Enter the path to your Excel file: ")
    
    if not os.path.exists(file_path):
        print("File not found!")
        return
    
    zero_percentage_students = process_excel_file(file_path)
    
    if zero_percentage_students is not None and len(zero_percentage_students) > 0:
        email_subject = "Important: Complete Your Classes - Action Required"
        email_template = """Dear {student_name},

We noticed that you haven't completed your classes yet (0% completion). 

Please log in to your account and complete the remaining classes as soon as possible. Your progress is important for your academic success.

If you have any questions or need assistance, please don't hesitate to contact us.

Best regards,
Academic Team"""
        
        success_count = 0
        error_count = 0
        
        for i, (_, student) in enumerate(zero_percentage_students.iterrows()):
            print(f"Processing {i+1}/{len(zero_percentage_students)}: {student['Name']}")
            
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
                print(f"✅ Email sent to {student['Name']} ({student['Email']})")
            else:
                error_count += 1
                print(f"❌ Failed to send email to {student['Name']}: {message}")
        
        print(f"\n🎉 Successfully sent {success_count} emails!")
        if error_count > 0:
            print(f"⚠️ Failed to send {error_count} emails.")
    else:
        print("No students with 0% completion found.")

if __name__ == "__main__":
    main()
