import seperator
import extractor
import mail
from time import sleep

def get_user_input():
    print("What would you like to do?"), sleep(0.3)
    print("    1. Put all submissions into one file")
    print("    2. Extract user IDs and send mails"), sleep(0.5)

    while True:
        user_input = input("Enter your choice (1 or 2): ")
        if user_input in ['1', '2']:
            return user_input
        else:
            print(f"{user_input} is an invalid choice. Please enter 1 or 2.")

def main():
    print("Starte Klausurenrückgabe...\n")
    sleep(1)

    user_input = get_user_input()
    if user_input == '1':
        extractor_input = input("\nExtract data? (y/n) ")
        if extractor_input == "y":
            extractor.main()
        elif extractor_input == "n":
            exit('Exiting...')

    elif user_input == '2':
        mail_input = input("\nSend all mails? (y/n): ")
        if mail_input == "y":
            mail.main()

        elif user_input == "n":
            exit("Exiting...")


if __name__ == "__main__":
    main()