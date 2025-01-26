import pandas as pd
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Function to send email with attachments
def send_email(recipient_email, subject, body, attachment_path):
    sender_email = "abc@gmail.com"  # Replace with your email
    sender_password = "abc def ghi jkl"  # Replace with your app password(It is not your gmail password)

    try:
        # Set up the MIME structure
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = recipient_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html"))  # Send email as HTML

        # Attach PNG file
        if os.path.exists(attachment_path):
            with open(attachment_path, "rb") as attachment_file:
                part = MIMEApplication(attachment_file.read(), Name=os.path.basename(attachment_path))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
                msg.attach(part)

        # Connect to the Gmail SMTP server
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, sender_password)
            server.send_message(msg)

        print(f"Email successfully sent to {recipient_email}")
        return True
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")
        return False

# Main function
if __name__ == "__main__":
    # Read the Excel file
    file_path = "students.xlsx"  # Your Excel file path
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Failed to read the file {file_path}: {e}")
        exit()

    # Check if the file has the necessary columns
    required_columns = ["Name", "Email"]
    for col in required_columns:
        if col not in df.columns:
            print(f"Missing column: {col} in the Excel file {file_path}")
            exit()

    # Process each student
    for index, row in df.iterrows():
        try:
            name = row["Name"]
            email = row["Email"]
            png_file = f"{name}.png"  # Assuming the PNG file is named after the student's name

            # Skip rows with missing essential information
            if not name or not email:
                print(f"Skipping row {index + 1}: Missing required data.")
                continue

            # Email subject and body
            subject = "Invitation to Yathartha Exhibition"
            body = f"""
            Dear {name},<br><br>
        We are pleased to invite you to the <b> Yathartha Exhibition</b> happening on <b>(Date here please).</b><br><br>
        Please find the invitation attached.<br><br>
        Best regards,<br>
        IOE, Thapathali Campus
            """

            # Send the email
            send_email(email, subject, body, png_file)

        except Exception as e:
            print(f"Error processing row {index + 1}: {e}")