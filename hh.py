import re
out = """send-mailmessage : Error in processing. The server response was: 4.7.0 Temporary server error. Please try again later. nPRX2
       At line:1 char:1
       + send-mailmessage -to tester@exch.local -subject ddd -body dss -smtpse ...
       + ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
       + CategoryInfo          : InvalidOperation: (System.Net.Mail.SmtpClient:SmtpClient) [Send-MailMessage], SmtpException
       + FullyQualifiedErrorId : SmtpException,Microsoft.PowerShell.Commands.SendMailMessage"""

expression = "PRX2"
if re.findall("PRX3", out):
    print("Possible solution:\n")
    print("Add IPs of this machine and DC to hosts file of Exchange machine where command executed.")