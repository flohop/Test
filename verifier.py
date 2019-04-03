import csv
import os

file_dir = os.path.dirname(os.path.abspath(__file__))
base_path = os.path.dirname(file_dir) + "\WebScraping"
dependencies = base_path + "\dependencies"


def check_verification(ver_mail):
    """Check if a certain email is in the .csv file (already verified)"""
    with open( dependencies + '\mail_verification.csv', 'r') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        in_file = False
        for row in csv_reader:
            try:
                if row[0] == str(ver_mail):
                    in_file = True
                    break
            except IndexError:
                break
        if in_file:
            print("Email already verified")
        else:
            print("Email not verified")


def verify(ver_mail, verif_code, given_code):
    """"If verification correct, append the mail to the list if verified mails"""
    with open(dependencies + '\mail_verification.csv', 'w') as csv_file:
        csv_writer = csv.writer(csv_file, delimiter=",")
        line_count = 0

        if verif_code == given_code:

            try:
                csv_writer.writerow([ver_mail, verif_code, "Verified"])
                print("Code correct")
            except PermissionError:
                print("Please close the file")
        else:
            print("Code not correct")


given_code = 1234
ver_code = 1234
check_verification("me@gmail.com")
verify("me@gmail.com", ver_code, given_code)


