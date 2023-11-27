dataDir = "Data/"

# Create MailMessage instance by loading an Eml file
message = MailMessage.load(dataDir + "test.eml")

# Get the sender info, recipient info, subject, html body and text body 
print("Sender: " + str(message.from_address))

for receiver in enumerate(message.to):
    print("Receiver: " + receiver)

print("Subject: " + message.subject)

print("HtmlBody: " + message.html_body)

print("TextBody: " + message.body)