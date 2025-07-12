#
#
#   Author: David Peprah
#   Description: This script updates send new students data to the building secretaries
#

import os, re
import sys
from configparser import ConfigParser
import logging, argparse, datetime, csv
from jinja2 import Environment, FileSystemLoader
from lib import send_email

# set up the Jinja2 environment to load templates from the 'templates' directory
env = Environment(loader=FileSystemLoader('templates'))


# Function to reset student password
def reset_student_password(username: str):
    """
    Function to reset a student's password.
    This function calls the PowerShell script to reset the password.
    """
    try:
        import subprocess
        # Call the PowerShell script with the username as an argument
        result = subprocess.run(['powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', 'lib\reset_password.ps1', '-username', username], capture_output=True, text=True)
        
        # Read Information from Powershell
        message = str(result.communicate()[0][:-2], 'utf-8')
        status, update = message.split('\r\n')

        if status == "Success":
            logging.info(f"Password reset successfully for student: {username}")
            return True
        else:
            logging.error(f"Failed to reset password for student {username}: {update}")
            return False
        
    except Exception as e:
        logging.error(f"An error occurred while resetting password: {e}")
        return False


def get_new_student_data():
    """
    Function to get new student data.
    This function can be implemented to fetch new student data from a source.
    """
    # get yesterday's date
    yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
    date = yesterday.strftime("%m-%d-%Y")

    if not args.testing:
        if not config.get('general', 'dataFolder'):
            logging.CRITICAL("Data folder path is not configured in the config file.")
            sys.exit(1)

        # check if the folder exists
        base_folder = config.get('general', 'dataFolder')
        if not os.path.exists(base_folder):
            logging.error(f"Folder does not exist: {base_folder}")
            sys.exit(1)

        
        abs_folder_path = f"{base_folder}\{date}\StudentCreated.csv"
        # Check if the file exists
        if not os.path.exists(abs_folder_path):
            logging.info(f"No new Student Created {abs_folder_path}")
            sys.exit(0)
    else:
        base_folder = r'config\sample_student_data'
        if not os.path.exists(base_folder):
            os.makedirs(f"{base_folder}\{date}")
        # For testing purposes, use a sample file path
        abs_folder_path = rf'{base_folder}\StudentCreated.csv'
        if not os.path.exists(abs_folder_path):
            logging.CRITICAL(f"Sample file does not exist: {abs_folder_path}")
            sys.exit(1)


    csv_headers = config.get('general', 'csvFileHeaders')
    if not csv_headers:
        logging.CRITICAL("CSV file headers are not configured in the config file.")
        sys.exit(1)

    csv_headers = csv_headers.split(',')


    # Read the data from the file
    with open(abs_folder_path, 'r', encoding='utf-8') as csv_file:
        reader = csv.DictReader(csv_file)
        # Skip the header row
        next(reader)
        data = [row for row in reader]

    if not data:
        logging.info(f"No new student data found in {abs_folder_path}")
        sys.exit(0)

    # Create a dictionary to hold students grouped by building
    # The key will be the building name and the value will be a list of students in that building
    students_building = {}
    for student in data:
        # Extract the building name from the row
        building_name = student['School Name'].strip()
        if building_name not in students_building:
            students_building[building_name] = []
        # Append the student data to the respective building list
        students_building[building_name].append(student)

    for building_name in students_building.keys():
        # export students into a csv file
        building_formatted_name = "".join(re.split('[^a-zA-Z0-9]+', building_name))
        output_file_name = f"{building_formatted_name}_students.csv"
        output_location = os.path.join(base_folder, date, output_file_name)

        try:
            with open(output_location, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=csv_headers)
                writer.writeheader()
                for student in students_building[building_name]:
                    writer.writerow(student)
            logging.info(f"Exported {len(students_building[building_name])} students to {output_location} successfully.")

            secretary_email = config.get('BuildingSecretariesEmails', building_name, fallback=adminEmail)
            # Send email notification to building secretaries with a summary of the students
            send_email_notification(recipient=secretary_email, 
                                    subject=f"New Students Created for {building_name} on {date}",
                                    file_path=os.path.join(base_folder, date),
                                    file_name=output_file_name,
                                    template_name='new_students_email_template.html',
                                    with_attachment=True,
                                    message=f"Please find attached the list of new students created for {building_name} on {date}.")    
        except Exception as e:
            logging.exception(f"Error writing to CSV file {output_location}: {e}")

        
def send_email_notification(recipient: str = None, subject: str = " ", file_path: str = None, file_name: str = None, template_name: str = None, with_attachment: bool = False, message: str = "TESTING EMAIL NOTIFICATION"):
    
    if recipient:

        send_email_message = None

        if with_attachment:
            if file_path and file_name and template_name:
               
                email_template = env.get_template(template_name)
                subject = subject
                rendered_email = email_template.render()

                # Function to send email notification
                logger.info("Sending email notification...")
            
                try:
                    send_email_message = send_email.sendMessage('me', send_email.CreateMessageWithAttachment(serviceAccount, recipient, 
                                                                                subject, rendered_email, file_dir=file_path, filename=file_name))
                  
                except Exception as e:
                    logger.exception(f"Failed to send email notification with attachment: {e}")
            
            else:
                logger.critical("File path, file name or template name not provided for email notification with attachment")
            
            
        else:
            try:
                send_email_message = send_email.sendMessage('me', send_email.CreateMessage(serviceAccount, recipient, 
                                                                                    subject, message))
            except Exception as e:
                logger.exception(f"Failed to send email notification: {e}")

        # Check if the email was sent successfully
        if send_email_message:
            if send_email_message[0] == 'success':
                logger.info(f"Email notification sent to with subject {subject}")
            else:
                logger.error(f"Failed to send email notification: {send_email_message[1]}")

    else:   
        logger.critical("No recipient email provided or email template, skipping email notification")



def main():
    
    logger = logging.getLogger(__name__)

    if (args.reset_password):
        logger.info(f'Resetting password for student: {args.username}')
        # Here you would add the logic to reset the student's password
        # For example, call a function to reset the password
        reset_student_password(args.username)
    else:
        get_new_student_data() # Placeholder for other functionalities




if __name__ == "__main__":
    # Load Config file
    config = ConfigParser()
    config.read(r'config\\update-students.ini')
    if not config.sections():
        logging.CRITICAL("Configuration file is empty or not found.")
        sys.exit(1)

    logLevel = config.get('logs', 'logLevel' , fallback='INFO')
    logType = config.get('logs', 'logType', fallback='FILE')
    logFile = config.get('logs', 'logFile', fallback='update-students.log')
    
    serviceAccount = config.get('admin', 'serviceAccEmail')


    sysadmin = config.get('admin', 'sysadmin', fallback='dpeprah@vartek.com')
    adminEmail = config.get('admin', 'adminEmail', fallback=sysadmin)


    parser = argparse.ArgumentParser()
    parser.add_argument('-lL', '--logLevel', default=None, type=str, help='Set the logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)')
    parser.add_argument('-lT', '--logType', default=None, type=str, help='Set the logging output (CONSOLE, FILE)')
    parser.add_argument('-lF', '--logFile', default=None, type=str, help='Set the logging file path')
    parser.add_argument('-rp', '--reset_password', action='store_true', help='Reset student passwords')
    parser.add_argument('-u', '--username', type=str, help='Username of the student to update')
    parser.add_argument('-t', '--testing', type=str, action='store_true', help='For testing purposes only, do not use in production')

    args = parser.parse_args()

    logLevel = args.logLevel.upper() if args.logLevel else logLevel.upper()
    logType = args.logType.upper() if args.logType else logType.upper()
    logFile = args.logFile if args.logFile else logFile
    if logType not in ['CONSOLE', 'FILE']:
        logging.CRITICAL('Invalid log type: %s' % logType)
        sys.exit(1)
    
    
    if args.reset_password:
        if not args.username:
            logging.CRITICAL('Username must be provided when resetting password.')
            sys.exit(1)
      

    numeric_level = getattr(logging, logLevel)
    if not isinstance(numeric_level, int):
        raise ValueError('Invalid log level: %s' % logLevel)

    if logType == 'FILE':
        logging.basicConfig(level=numeric_level, 
                            format="{asctime} - {levelname} - {message}", 
                            style="{", 
                            filemode='a+', 
                            filename=logFile,
                            encoding="utf-8",
                            datefmt="%Y-%m-%d %H:%M:%S")
    elif logType == 'CONSOLE':
        logging.basicConfig(level=numeric_level, 
                            format="{asctime} - {levelname} - {message}", 
                            style="{", 
                            encoding="utf-8",
                            datefmt="%Y-%m-%d %H:%M")
    logger = logging.getLogger(__name__)

    main()
