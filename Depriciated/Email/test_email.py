from email_handler import send_email

result = send_email(
    to_email="shaikhaunraza@gmail.com",
    subject="ProVA Test Email",
    message_text="Hello ! This is a test email from the ProVA email module.",
)

print(result)