import smtplib
import ssl

# User configuration
sender_email = "thecallcentre4@gmail.com"
receiver_emails = ['m.t.atheros@gmail.com', 'info@thecallcentre.co.nz']#, 'john@thecallcentre.co.nz']
password = "Naz55105"

coupon_count = 2
# Email text
email_body = '\n\nWILSONS COUPONS: x %s' % coupon_count


# Creating a SMTP session | use 587 with TLS, 465 SSL and 25
server = smtplib.SMTP('smtp.gmail.com', 587)
# Encrypts the email
context = ssl.create_default_context()
server.starttls(context=context)
# We log in into our Google account
server.login(sender_email, password)
# Sending email from sender, to receiver with the email body
server.sendmail(sender_email, receiver_emails, email_body)
server.quit()
