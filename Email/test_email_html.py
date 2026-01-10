from email_handler import send_email

send_email(
    to_email="nidaqureshin04@gmail.com",
    subject="HTML Test",
    message_text="<h2>Hello from ProVA</h2>",
    is_html=True
)
