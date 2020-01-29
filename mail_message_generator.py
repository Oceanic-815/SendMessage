"""
Script for generating email messages within one mailbox in MS Exchange environment.
Options described below
"""
# TODO - create ability to specify '--file_size' option to generate a file for attachment
import subprocess
import optparse
import sys

counter = 1  # used for printing how many messages created


def filter_none(option):  # To avoid concatenation None type to string we filter options from None values
    if option is not None:
        return option


def generator():
    #  Creation of options
    parser = optparse.OptionParser()
    parser.add_option("-t", "--to", type="string", help="Target mailbot, e.g target@test.local")
    parser.add_option("-f", "--from", dest="fr", type="string", help="Sender mailbox, e.g sender@test.local (or any)")
    parser.add_option("-s", "--subject", dest="sbj", type="string", help="Enter any text for subject")
    parser.add_option("-a", "--attachment", dest="att", type="string", help=r"Enter a path to a file, e.g. C:\file.txt")
    parser.add_option("-r", "--smtp", dest="smtpserv", type="string",
                      help="Enter SMTP server name or address, e.g server.test.local")
    parser.add_option("-c", "--count", dest="cnt", type="int", help="Specify number how many messages to send, e.g 10")
    opts, args = parser.parse_args()  # Parsing options
    send_to = opts.to  # assign specified options to variable. opts.to is the value under --to option
    send_from = opts.fr
    subject = opts.sbj
    attachments = opts.att
    smtpserver = opts.smtpserv
    count = opts.cnt
    to_resulted = " "  # Variables to create a PowerShell '--parameter value' things
    from_resulted = " "
    subject_resulted = " "
    attachments_resulted = " "
    smtpserver_resulted = " "
    count_resulted = 1
    #  Try-Excepts to catch exceptions on assigning values to the options and make user-friendly errors
    try:
        to_resulted = filter_none("-to " + send_to + " ")
    except Exception:
        print("'-to' option not specified, continue...")
    try:
        from_resulted = filter_none("-from " + send_from + " ")
    except Exception:
        print("'-from' option not specified, continue...")
    try:
        subject_resulted = filter_none("-subject " + subject + " ")
    except Exception:
        print("'-subject' not specified, continue...")
    try:
        attachments_resulted = filter_none("-attachments " + attachments + " ")
    except Exception:
        print("'-attachments' not specified, continue...")
    try:
        smtpserver_resulted = filter_none("-smtpserver " + smtpserver + " ")
    except Exception:
        print("'-smtpserver' not specified, continue...")
    try:
        count_resulted = filter_none(count)
    except Exception:
        print("'-count' not specified, continue...")
    try:
        for i in range(count_resulted):  # Executing command '--count' number of times
            try:
                gen = subprocess.run(
                    ["Powershell", "send-mailmessage", to_resulted, from_resulted, subject_resulted,
                     attachments_resulted, smtpserver_resulted], timeout=30, check=True, stdout=subprocess.PIPE)
                result = gen.stdout.decode('utf-8')
                print(result)
            except subprocess.CalledProcessError as e:
                print(e.output, "\n\n!!! No messages have been sent due to above exception !!!\n")
                sys.exit()
            print("=== ", counter + i, " message(s) created")
    except TypeError as err:
        print("\n", err, "\nCheck if --count parameter is specified")


if __name__ == '__main__':
    try:
        generator()
    except KeyboardInterrupt:
        print("\nOperation canceled by user")
    except subprocess.TimeoutExpired:
        print("Timeout error")
    except Exception:
        raise
