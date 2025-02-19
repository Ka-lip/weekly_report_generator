# v1.0

from secret import *
import fetch_mail
import weekly

def main():
    fetch_mail.fetch_outlook_emails()
    weekly.main()

if __name__ == "__main__":
    main()