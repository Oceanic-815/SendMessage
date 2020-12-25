"""
Script for generating email messages within one mailbox in MS Exchange environment.
All options have to be specified with values to populate a specified mailbox with emails.
Emails can be sent with attachments or without them. Attachment should be specified as a size in bytes without quotes.
If attachment option is specified, a file with specified size will be created in %TEMP% directory.
This file will be attached to each email message and automatically deleted after script completion.
Make sure that the size of the file is below the limits of Exchange/mailbox settings. Otherwise, script will fail.
Available options described below.
"""
#  TODO  need to create an option that increase mailbox and DB size limits
import subprocess
import optparse
import random
import sys
import os
import re

counter = 1  # used for printing how many messages created
path_to_temp_file = os.environ.get("TEMP") + "\\temp.file"  # temp file is created in %TMP% folder
set_subject = 'abcdefghigklmnopqrstuvwxyz1234567890'


def rand_subject(lenght):
    return ''.join(random.choice(set_subject) for i in range(lenght))


def generate_file(size):  # Generate a file to attach to a message
    with open(path_to_temp_file, 'wb') as resulted_file:
        resulted_file.write(size)
        resulted_file.close()
        print("\nFile " + path_to_temp_file + " created for attachments\n")


def filter_none(option):  # To avoid concatenation None type to string we filter options from None values
    if option is not None:
        return option


# Increase limits Set-TransportConfig and Set-ReceiveConnector for current host
def increase_limits(*args, **kwargs):
    get_comp_name = ('\'' + "Default " + os.environ['COMPUTERNAME'] + '\'')
    cmd = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -ExecutionPolicy RemoteSigned -noexit -command \". 'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1'; Connect-ExchangeServer -auto; Set-TransportConfig -MaxSendSize unlimited; Set-TransportConfig -MaxReceiveSize unlimited; Set-ReceiveConnector " + get_comp_name + " -MaxMessageSize '1.9GB'; Set-SendConnector " + get_comp_name + " -MaxMessageSize '1.9GB' \""
    com = subprocess.run(cmd, stdout=subprocess.PIPE)
    # print(cmd)
    res = com.stdout.decode('utf-8')
    print(res)


# To set –MaxSendSize and –MaxReceiveSize for each mailbox, we need to get a list mailboxes. This func is doing it
def getting_mailboxes():
    get_mailbox_cmd = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -ExecutionPolicy RemoteSigned -noexit -command \". 'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1'; Connect-ExchangeServer -auto; Get-Mailbox | fl name \""
    get_mailbox_exec = subprocess.run(get_mailbox_cmd, stdout=subprocess.PIPE)
    get_mailbox_raw_output = get_mailbox_exec.stdout.decode('utf-8')
    get_mailbox_splitted_output = get_mailbox_raw_output.split("\n")
    get_mailbox_raw_list = []
    for i in get_mailbox_splitted_output:
        if "Name : " in i:
            get_mailbox_raw_list.append(i)
    list_of_names = []
    for i2 in get_mailbox_raw_list:
        list_of_names.append(i2[7:].replace("\r",""))
    return list_of_names


# May be useful to set –MaxSendSize and –MaxReceiveSize for each mailbox
def setting_mailbox_rx_size_limits():
    for item in getting_mailboxes():
        set_command = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -ExecutionPolicy RemoteSigned -noexit -command \". 'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1'; Connect-ExchangeServer -auto; Set-Mailbox " + item + " –MaxSendSize 1024MB –MaxReceiveSize 1024MB \""
        run_set_mailbox = subprocess.run(set_command, stdout=subprocess.PIPE)
        set_mailbox_output = run_set_mailbox.stdout.decode('utf-8')
        print(set_mailbox_output)


def generator():
    #  Creation of options
    parser = optparse.OptionParser()
    parser.add_option("-t", "--to", type="string", help="Target mailbot, e.g target@test.local")
    #parser.add_option("-f", "--from", dest="fr", type="string", help="Sender mailbox, e.g sender@test.local (or any)")
    parser.add_option("-a", "--attachment_size", dest="att", type="int",
                      help=r"Enter size for file to attach (in bytes)")
    parser.add_option("-m", "--smtp", dest="smtpserv", type="string",
                      help="Enter SMTP server name or address, e.g server.test.local")
    parser.add_option("-c", "--count", dest="cnt", type="int", help="Specify number how many messages to send, e.g 10")
    parser.add_option("-i", "--increaselimits", dest="increase", action="callback", callback=increase_limits, help="Increases messages/attachment size limits on Exchange")
    opts, args = parser.parse_args()  # Parsing options
    send_to = opts.to  # assign specified options to variable. opts.to is the value under --to option
    # send_from = opts.fr # Used before when --from option was necessary
    if opts.att is not None:
        generate_file(os.urandom(opts.att))  # size in bytes specified in CMD option --attachment_size or -a
    smtpserver = opts.smtpserv
    count = opts.cnt
    to_resulted = " "  # Variables to create a PowerShell '--parameter value' things
    from_resulted = " "
    attachments_resulted = " "
    smtpserver_resulted = " "
    count_resulted = 1
    #  Try-Excepts to catch exceptions on assigning values to the options and make user-friendly errors
    try:
        to_resulted = filter_none("-to " + send_to + " ")
    except Exception:
        pass
    try:
        from_resulted = filter_none("-from " + send_to + " ")  # send_from was changed to send_to. send_from not needed
    except Exception:
        pass
    try:
        if opts.att is None:  # Checking if --attachment_size option not specified, specify this as empty string
            attachments_resulted = ""
        else:
            attachments_resulted = filter_none("-attachments " + path_to_temp_file + " ")
    except Exception:
        print("'-attachments' not specified, continue...")
    try:
        smtpserver_resulted = filter_none("-smtpserver " + smtpserver + " ")
    except Exception:
        pass
    try:
        count_resulted = filter_none(count)
    except Exception:
        print("\n'-count' should be specified!\n")

    try:
        for i in range(count_resulted):  # Executing command '--count' number of times
            subject_resulted = str(rand_subject(50))
            try:
                gen = subprocess.run(
                    ["Powershell", "send-mailmessage", to_resulted, from_resulted, subject_resulted,
                     attachments_resulted, smtpserver_resulted], timeout=60, check=True, stdout=subprocess.PIPE)
                result = gen.stdout.decode('utf-8')
                print(result)
            except subprocess.CalledProcessError as e:
                print(e.output, "\n\n!!! The message has not been sent due to above exception !!!\n")
                try:
                    os.remove(path_to_temp_file)
                except Exception:
                    print("\n")
                if re.findall("PRX2", str(e.output)):
                    print("POSSIBLE SOLUTION:")
                    print("Add IPs of this machine and DC to 'hosts' file of Exchange machine where command executed.")
                sys.exit()
            print("=== ", counter + i, " message(s) created")
    except TypeError as err:
        print("\n", err, "\nCheck if --count parameter is specified. \nType -h for help")


if __name__ == '__main__':
    try:
        generator()
    except KeyboardInterrupt:
        print("\nOperation canceled by user\n")
    except subprocess.TimeoutExpired:
        print("Timeout error")
    except Exception:
        raise
    try:
        os.remove(path_to_temp_file)
    except Exception:
        print("\n")

# pyinstaller --onefile -i C:\Projects\SendMessage\mail_message_generator.ico --noupx mail_message_generator.py


#C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit -command ". 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell "

# Set-TransportConfig -MaxSendSize unlimited
# Set-TransportConfig -MaxReceiveSize unlimited

# ??????????
# Set-ReceiveConnector "Default WIN-OL" -MaxMessageSize 1.9GB
# Name                           MaxMessageSize
# ----                           --------------
# Default WIN-OL                 1.9 GB (2,040,109,465 bytes)
# Client Proxy WIN-OL            1.9 GB (2,040,109,465 bytes)
# Default Frontend WIN-OL        1.9 GB (2,040,109,465 bytes)
# Outbound Proxy Frontend WIN-OL 1.9 GB (2,040,109,465 bytes)
# Client Frontend WIN-OL         1.9 GB (2,040,109,465 bytes)
