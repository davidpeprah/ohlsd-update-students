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
def reset_student_password(username: str, building: str):
    """
    Function to reset a student's password.
    This function calls the PowerShell script to reset the password.
    """
    building_short_names = {
        'RRMS': 'Rapid Run Middle School', 
        'TDS': 'Test Dummy School',
        'SPG': 'Charles W. Springmyer Elementary',
        'OHHS': 'Oak Hills High School',
        'JFD': 'John Foster Dulles Elementary',
        'COH': 'C.O. Harrison Elementary',
        'DEL': 'Delshire Elementary',
        'OAK': 'Oakdale Elementary',
        'BMS': 'Bridgetown Middle School',
        'DMS': 'Delhi Middle School'
    }
    
    try:
        import subprocess

        cc = adminEmail


        if args.testing:
            # For testing purposes, simulate the password reset
            username = 'test_user'
            logger.debug(f"Simulating password reset for student: {username}")
            building = 'TDS'
            cc = sysadmin


        logger.info(f"Resetting password for student: {username}")

        # Call the PowerShell script with the username as an argument
        result = subprocess.Popen(['powershell.exe', '-ExecutionPolicy', 
                                   'Bypass', '-File', r'lib\\reset_password.ps1', 
                                   '-username', username], stdout=subprocess.PIPE)
        
        # Read Information from Powershell
        message = str(result.communicate()[0][:-2], 'utf-8')
        status, update = message.split('\r\n')
        logger.debug(f"PowerShell script output: {message}")

        if status == "Success":
            logger.info(f"Password reset successfully for student: {username}")
            displayName, password = update.split(',')
           
            # get building secretary email from config
            building_name = building_short_names.get(building.upper(), None)
            secretary_email = config.get('BuildingSecretariesEmails', building_name, fallback=adminEmail)
            
            # get data for email notification
            data = {"Fullname": displayName.strip(), "Username": username, "Password": password.strip(), "building": building_name}
            
            # Send email notification to the building secretary
            send_email_notification(data=data, 
                                    recipient=secretary_email,
                                    subject=f"Password Reset Notification for {displayName.strip()}",
                                    template_name='password_reset_email_template.html',
                                    cc=cc)
        else:
                logger.error(f"Failed to reset password for student {username}: {update}")
            
        
    except Exception as e:
        logger.error(f"An error occurred while resetting password: {e}")
    


def get_new_student_data():
    """
    Function to get new student data.
    This function can be implemented to fetch new student data from a source.
    """
    # get yesterday's date
    yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
    date = yesterday.strftime("%m-%d-%Y")
    cc = adminEmail

    if not args.testing:
        if not config.get('general', 'dataFolder'):
            logger.CRITICAL("Data folder path is not configured in the config file.")
            sys.exit(1)

        # check if the folder exists
        base_folder = config.get('general', 'dataFolder')
        if not os.path.exists(base_folder):
            logger.error(f"Folder does not exist: {base_folder}")
            sys.exit(1)

        
        abs_folder_path = rf"{base_folder}\{date}\StudentCreated.csv"
        # Check if the file exists
        if not os.path.exists(abs_folder_path):
            logger.info(f"No new Student Created {abs_folder_path}")
            sys.exit(0)
    else:
        base_folder = r'config\sample_student_data'
        if not os.path.exists(rf"{base_folder}\{date}"):
            logger.debug(f"Creating folder for date: {date}")
            os.mkdir(rf"{base_folder}\{date}")
        # For testing purposes, use a sample file path
        abs_folder_path = rf'{base_folder}\StudentCreated.csv'
        if not os.path.exists(abs_folder_path):
            logger.CRITICAL(f"Sample file does not exist: {abs_folder_path}")
            sys.exit(1)
        cc = sysadmin


    csv_headers = config.get('general', 'csvFileHeaders')
    if not csv_headers:
        logger.CRITICAL("CSV file headers are not configured in the config file.")
        sys.exit(1)

    csv_headers = csv_headers.split(',')
    logger.debug(f"CSV Headers: {csv_headers}")

    # Read the data from the file
    with open(abs_folder_path, 'r', encoding='utf-8') as csv_file:
        logger.debug(f"Reading data from {abs_folder_path}")
        reader = csv.DictReader(csv_file)
        # Skip the header row
        next(reader)
        data = [row for row in reader]

    if not data:
        logger.info(f"No new student data found in {abs_folder_path}")
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
            logger.info(f"Exported {len(students_building[building_name])} students to {output_location} successfully.")

            secretary_email = config.get('BuildingSecretariesEmails', building_name, fallback=adminEmail)
            logger.debug(f"Building: {building_name}, Secretary Email: {secretary_email}")
            if not secretary_email:
                logger.error(f"No email configured for building: {building_name}. Skipping email notification.")
                continue

            # remove any new line characters from the email address
            secretary_email = [email.strip() for email in secretary_email.split(',')]
            # join the emails back with comma separator
            secretary_email = ','.join(secretary_email)
                
            # Send email notification to building secretaries with a summary of the students
            send_email_notification(data={"building": building_name, "date": date, "students_count": len(students_building[building_name])},
                                    recipient=secretary_email, 
                                    subject=f"New Students Created for {building_name} on {date}",
                                    file_path=os.path.join(base_folder, date),
                                    file_name=output_file_name,
                                    template_name='new_students_email_template.html',
                                    with_attachment=True,
                                    cc=cc)    
        except Exception as e:
            logger.exception(f"Error writing to CSV file {output_location}: {e}")

        
def send_email_notification(data: dict = None, recipient: str = None, subject: str = " ", file_path: str = None, file_name: str = None, template_name: str = None, with_attachment: bool = False, message: str = "TESTING EMAIL NOTIFICATION", cc: str = None):
    
    if recipient:

        send_email_message = None

        if with_attachment:
            if file_path and file_name and template_name:
               
                email_template = env.get_template(template_name)
                rendered_email = email_template.render(data)

                # Function to send email notification
                logger.info(f"Sending email notification with attachment subject: {subject} ...")
                logger.debug(f"Email subject: {subject}")
                logger.debug(f"Email recipient: {recipient}")
                try:
                    send_email_message = send_email.sendMessage('me', send_email.CreateMessageWithAttachment(serviceAccount, recipient, 
                                                                                subject, rendered_email, file_dir=file_path, filename=file_name, cc=cc))
                  
                except Exception as e:
                    logger.exception(f"Failed to send email notification with attachment: {e}")
            
            else:
                logger.critical("File path, file name or template name not provided for email notification with attachment")   
        else:
            if not template_name:
                logger.critical("Email template name not provided for email notification without attachment")
                return
            
            email_template = env.get_template(template_name)
            rendered_email = email_template.render(data)

            # Function to send email notification
            logger.info(f"Sending email notification without attachment subject: {subject} ...")
            logger.debug(f"Email subject: {subject}")
            logger.debug(f"Email recipient: {recipient}")
            try:
                send_email_message = send_email.sendMessage('me', send_email.CreateMessageWithAttachment(serviceAccount, recipient, 
                                                                                    subject, rendered_email))
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
        reset_student_password(args.username, args.building.upper())
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
    logFile = config.get('logs', 'logFile', fallback=r'logs\\update-students.log')
    
    serviceAccount = config.get('admin', 'serviceAccEmail')


    sysadmin = config.get('admin', 'sysadmin', fallback='dpeprah@vartek.com')
    adminEmail = config.get('admin', 'adminEmail', fallback=sysadmin)


    parser = argparse.ArgumentParser()
    parser.add_argument('-lL', '--logLevel', default=None, type=str, help='Set the logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)')
    parser.add_argument('-lT', '--logType', default=None, type=str, help='Set the logging output (CONSOLE, FILE)')
    parser.add_argument('-lF', '--logFile', default=None, type=str, help='Set the logging file path')
    parser.add_argument('-rp', '--reset_password', action='store_true', help='Reset student passwords')
    parser.add_argument('-u', '--username', type=str, help='Username of the student to update')
    parser.add_argument('-b', '--building', type=str, help='Student building name for email notification use => [RRMS, TDS, SPG, OHHS, JFD, COH, DEL, OAK, BMS, DMS]')   
    parser.add_argument('-t', '--testing', action='store_true', help='For testing purposes only, do not use in production')

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

        if args.building:
            if not re.match(r'^[a-zA-Z\s]+$', args.building):
                logging.CRITICAL('Building name must contain only letters and spaces.')
                sys.exit(1)
            
            if args.building.upper() not in ['RRMS', 'TDS', 'SPG', 'OHHS', 'JFD', 'COH', 'DEL', 'OAK', 'BMS', 'DMS']:
                logging.CRITICAL(f"Invalid building name: {args.building.upper()}. Must be one of: RRMS, TDS, SPG, OHHS, JFD, COH, DEL, OAK, BMS, DMS")
                sys.exit(1)
        

    numeric_level = getattr(logging, logLevel)
    if not isinstance(numeric_level, int):
        logging.CRITICAL(f"Invalid log level: {logLevel}")
        sys.exit(1)

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
