import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os

# פונקציה לשליחת מייל
def send_email(recipient_email, file_path):
    sender_email = "b4107426@gmail.com"

    try:
        # יצירת המייל
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = "Your Processed Excel File"
        body = "Attached is your processed Excel file."
        msg.attach(MIMEBase('text', 'plain'))
        
        # הוספת הקובץ למייל
        with open(file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={os.path.basename(file_path)}",
        )
        msg.attach(part)

        # שליחת המייל
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
        return "Email sent successfully!"
    except Exception as e:
        return f"Failed to send email: {e}"

# ממשק משתמש
st.title("AI Excel Assistant")
st.write("Upload your Excel file, and I'll process it for you.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # קריאת הטבלה
        df = pd.read_excel(uploaded_file)
        st.write("Here's your data:")
        st.dataframe(df)

        # הוספת שורת סה"כ רק אם יש עמודות מספריות
        if st.button("Add Total Row"):
            numeric_columns = df.select_dtypes(include='number')
            if not numeric_columns.empty:
                total_row = numeric_columns.sum()
                total_row["Total"] = "Total"
                df = df.append(total_row, ignore_index=True)
                st.write("Here's the updated data with Total Row:")
                st.dataframe(df)
                
                # שמירת הקובץ
                output_file = "processed_file_with_totals.xlsx"
                df.to_excel(output_file, index=False)
                
                # שואל את המשתמש לאיזה מייל לשלוח את הקובץ
                recipient_email = st.text_input("Enter your email to receive the processed file:")

                if recipient_email:
                    result = send_email(recipient_email, output_file)
                    st.success(result)
                else:
                    st.warning("Please enter a valid email address.")
            else:
                st.warning("No numeric columns found. Total row not added.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
