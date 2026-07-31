import os
import smtplib
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


class Email:
    def send_email(self, sender_email, password, receiver_email, subject, body, filenames):
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = receiver_email
        msg["Subject"] = subject

        fixed_cc = "boris.wang@deltaww.com"
        msg["Cc"] = fixed_cc
        msg.attach(MIMEText(body, "html"))

        if filenames is None:
            filenames = []
        elif isinstance(filenames, str):
            filenames = [filenames]

        for filename in filenames:
            if not os.path.isfile(filename):
                print(f"File {filename} does not exist.")
                continue

            with open(filename, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f'attachment; filename="{os.path.basename(filename)}"',
                )
                msg.attach(part)

        smtp_server = "deltarelay.deltaww.com"
        smtp_port = 25
        smtp_timeout = 120
        receiver_list = []
        receiver_keys = set()
        for receiver in [r.strip() for r in receiver_email.split(",") if r.strip()] + [fixed_cc]:
            receiver_key = receiver.lower()
            if receiver_key not in receiver_keys:
                receiver_list.append(receiver)
                receiver_keys.add(receiver_key)

        last_error = None
        for attempt in range(1, 3):
            server = None
            try:
                server = smtplib.SMTP(smtp_server, smtp_port, timeout=smtp_timeout)
                server.ehlo()

                if server.has_extn("starttls"):
                    server.starttls()
                    server.ehlo()

                if password and server.has_extn("auth"):
                    server.login(sender_email, password)

                server.sendmail(sender_email, receiver_list, msg.as_string())
                print("Email sent successfully!")
                return

            except smtplib.SMTPAuthenticationError as auth_e:
                print(
                    "Failed to send email. Error: SMTP Authentication failed. "
                    f"Check username and password in config. Error: {auth_e}"
                )
                raise
            except smtplib.SMTPSenderRefused as sender_e:
                if sender_e.smtp_code == 530 and not password:
                    print(
                        "Failed to send email. This client is not allowed to send anonymous mail. "
                        "Set DELTA_SMTP_PASSWORD for SMTP authentication, or ask IT to allow this "
                        "PC/IP to use deltarelay anonymously."
                    )
                print(f"Failed to send email. Sender was refused by SMTP server: {sender_e}")
                raise
            except (TimeoutError, smtplib.SMTPServerDisconnected) as e:
                last_error = e
                print(f"Email send attempt {attempt} failed. Error: {e}")
                if attempt == 2:
                    raise
                time.sleep(5)
            except Exception as e:
                print(f"Failed to send email. Error: {e}")
                raise
            finally:
                if server:
                    try:
                        server.quit()
                    except smtplib.SMTPServerDisconnected:
                        pass

        if last_error:
            raise last_error
