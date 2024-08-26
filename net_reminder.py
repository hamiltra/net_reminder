"""

Script: net_reminder.py
Purpose: To send an HTML email weekly based on a user-defined template
Author: Roger Hamilton, KK6LZB

"""
# Date handling
from datetime import datetime, timedelta
# Import the email modules we'll need
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
# Command line args
import getopt
import logging
from logging import handlers
import os
import sys
# Import smtplib for the actual sending function
import smtplib
# Email templating
from jinja2 import Template
# Easy handling of data
import pandas as pd
from pandas.core.groupby.groupby import DataError
# Configuration file
import yaml

#######################################
# Declaring global vars
#######################################
SCRIPT_CONFIG = None
OPTS = None
TEMPLATE = None

#######################################
# Command Line Options
#######################################
LOG_NAME = None
CONFIG_FILE = None
EMAIL_CONFIG = None
NO_NET_CONTROL_EMAIL_CONFIG = None
NOW = None
TEST = False
EMAIL_SUBJECT_TEMPLATE = None
NO_NET_CONTROL_EMAIL_SUBJECT_TEMPLATE = None
TEST_EMAIL = None
EMAIL_REPLY_TO = None

#######################################
# Sample email template
#######################################
DEFAULT_EMAIL_TEMPLATE = '''\
<!DOCTYPE html>
<head>
<style>
td {
  vertical-align: top;
}

.font-medium {
  font-size: 18px;
}
.font-small {
  font-size: small;
  color: Silver;
}
</style>
{% block title %}<h2>Amateur Radio {{ net_type }} Net for {{ net_date }}</h2> {% endblock %}
</head>
<body>
{% block content %}
<table>
<tr>
<td><img src="cid:{{logo}}"></td>
<td>
<p><h3>Amateur Radio {{ net_type }} Net Control Operators for {{net_date }}:</h3>
  <font class="font-medium">
  <ul>
      <li><b>Primary Net Control:</b> {{ primary_net_control }}
      <li><b>Backup Net Control:</b> {{ backup_net_control}}
  </ul>
  </font>
  <br>
  <b>HEADS UP:</b> For {{net_date_2wk}}, <b>Primary is {{ primary_net_control_2wk}}</b> and <b>Backup is {{ backup_net_control_2wk}}</b>.<br><br>
  
  <h4>Net Preparation:</h4>
    <ul>
      <li>For the Weekly Net, please arrive well in advance of 1900 hrs to allow time to setup and make the 1855 hr announcement. 
      <li>If you are scheduled for the Travel Net, then arrive well in advance of the 1825 hr announcement and 1830 hr Travel Net.
  </ul>
  <br>
  <b>NOTE:</b> If you are unable to perform as Primary Net Control, be sure to coordinate with Backup Net Control to ensure that the Net is covered. If Backup Net Control is <u>ALSO unavailable</u> , THE PRIMARY NET CONTROL must find a replacement and notify <a href="mailto:{{switch_notify1_email}}">{{switch_notify1_name}}, {{switch_notify1_email}}</a> or <a href="mailto:{{switch_notify2_email}}">{{switch_notify2_name}}, {{switch_notify2_email}}</a> of the switch.
  <br><br>
</p>
</td>
</tr>
<tr><td align='center' colspan='2'>For maintenance to the Net schedule, please contact {{excel_maintainer_name}} <a href='mailto: {{excel_maintainer_email}}'>{{excel_maintainer_email}}</a><br><br>
<font align='center' class="font-small">This notice automated using the net_reminder script <br>maintained by {{script_maintainer_name}}, <a href='mailto: {{script_maintainer_email}}'>{{script_maintainer_email}}</a></font><br>
</td></tr>
</table>
{% endblock %}
</body>
</html>
'''

DEFAULT_NO_NET_CONTROL_EMAIL_CONFIG = '''\
<!DOCTYPE html>
<head>
<style>
td {
  vertical-align: top;
}

.font-medium {
  font-size: 18px;
}
.font-small {
  font-size: small;
  color: Silver;
}
</style>
{% block title %}<h2>ATTENTION: NO On-call Net Control Operators Found</h2> {% endblock %}
</head>
<body>
{% block content %}
<table>
<tr>
<td>
<p><h3>No amateur radio Net Control Operators for {{net_date }} were found in the Excel configuration!!</h3>
</td>
</tr>
</table>
{% endblock %}
</body>
</html>
'''

#######################################
# Subroutines and Functions
#######################################
# Usage statement
def usage():
    """ Shows script usage 

        Args: None
        Returns: None
    """
    print("net_reminder.py - Net Reminder\n")
    print(" -h, --help        This notice.")
    print(" -c, --config      Configuration file. default: net_reminder.yaml")
    print(" -e, --econfig     Email layout configuration file. default: net_reminder.html")
    print(" -l, --log         Log file. default: net_reminder.log")
    print(" -n, --now         Now: mm/dd/yyyy. default: current date")
    print(" -s, --subject     Email subject. default: '<0> Net for <1>'")
    print(" -t, --test        Run script but don't send mail and print output. default: False")
    print(" -q, --test_email  Run script using test emails provided. default: None")
    print("")


def fill_email_no_net_notice_template():
    """Completes the email template
    
        Args: None

        Return: STRING email_body
    """

    email_template = None

        # Get the email configuration file from HTML template in file
    with open(NO_NET_CONTROL_EMAIL_CONFIG, 'r', encoding='UTF-8') as email_file_template:
        try:
            logging.info("Using email configuration file: %s", NO_NET_CONTROL_EMAIL_CONFIG)
            email_template = email_file_template.read()
            email_file_template.close()
        except IOError as exc:
            logger.fatal(exc)
            email_template = DEFAULT_NO_NET_CONTROL_EMAIL_CONFIG
        finally:
            email_file_template.close()

    #
    # HTML email template
    #
    t = Template(email_template)
    email_body = t.render(net_date=NOW.strftime("%m/%d/%Y"),
                          excel_maintainer_name=SCRIPT_CONFIG['excel_maintainer_name'],
                          excel_maintainer_email=SCRIPT_CONFIG['excel_maintainer_email'],
                          script_maintainer_name=SCRIPT_CONFIG['script_maintainer_name'],
                          script_maintainer_email=SCRIPT_CONFIG['script_maintainer_email'],
                          )

    return email_body

def create_no_net_email_subject():
    """Completes the email subject template and returns it
    
    Args: None

    Returns: STRING email_subject
    
    """
    net_date = NOW.strftime("%m/%d/%Y")
    email_subject = NO_NET_CONTROL_EMAIL_SUBJECT_TEMPLATE.format(net_date)
    logging.info("Email Subject: %s", email_subject)

    return email_subject


def gather_no_net_email_addresses():
    """Generates the email address list from defined maintainers' emails
    
    Args: None
    
    Returns: STRING email_dist: Comma-separated list of emails
    
    """
    email_dist = list([SCRIPT_CONFIG['excel_maintainer_email'],
                       SCRIPT_CONFIG['script_maintainer_email']])
    logging.info("Email Distribution List: %s", ",".join(email_dist),
                 )

    return email_dist

def email_no_net_notice ():
    """Sends a message to the maintainers that no net controls are present given the date

        Args: None

        Returns: None
    """
    email_body = fill_email_no_net_notice_template()
    email_subject = create_no_net_email_subject()
    email_dist = gather_no_net_email_addresses()

    msg = MIMEMultipart()
    text = MIMEText(email_body, 'html')
    msg.attach(text)

    me = SCRIPT_CONFIG['smtp_auth_user']
    logging.info("Email From: %s", me)

    if TEST is False:
        msg['Subject'] = email_subject
        msg['From'] = SCRIPT_CONFIG['email_from']
        msg['X-Priority'] = '2'

        if TEST_EMAIL is not None:
            logging.info("Sending test email")
            msg['To'] = TEST_EMAIL
            email_dist = [TEST_EMAIL]
        else:
            msg['To'] = ", ".join(email_dist)

        logger.info("Subject: %s", msg['Subject'])
        logger.info("From: %s", msg['From'])
        logger.info("To: %s", msg['To'])
        logger.info("CC: %s", msg['Cc'])
        logger.info("BCC: %s", msg['Bcc'])

        # Send the message via our own SMTP server, but don't include the
        # envelope header.
        s = smtplib.SMTP_SSL(SCRIPT_CONFIG['smtp_server'],
                             port=SCRIPT_CONFIG['smtp_port'])
        s.ehlo()
        s.login(me, SCRIPT_CONFIG['smtp_auth_pass'])
        s.sendmail(me, email_dist, msg.as_string())
        s.close()
    else:
        print(email_body)


def fill_email_net_notice_template(net_vars):
    """Completes the email template
    
        Args: None

        Return: STRING email_body
    """
    email_template = None
    logo_basename = None

    if net_vars['logo'] is not None:
        logo_basename = os.path.basename(net_vars['logo'])

    # Get the email configuration file from HTML template in file
    with open(EMAIL_CONFIG, 'r', encoding='UTF-8') as email_file_template:
        try:
            logging.info("Using email configuration file: %s", EMAIL_CONFIG)
            email_template = email_file_template.read()
            email_file_template.close()
        except IOError as exc:
            logger.warning(exc)
            email_template = DEFAULT_EMAIL_TEMPLATE
        finally:
            email_file_template.close()
    #
    # HTML email template
    #
    t = Template(email_template)
    email_body = t.render(net_date=net_vars['net_date'],
                          primary_net_control=net_vars['current_primary'],
                          backup_net_control=net_vars['current_backup'],
                          net_type=net_vars['net_type'],
                          logo=logo_basename,
                          primary_net_control_2wk=net_vars['next_primary'],
                          backup_net_control_2wk=net_vars['next_backup'],
                          net_date_2wk=net_vars['next_net_date'],
                          switch_notify1_name=SCRIPT_CONFIG['switch_notify1_name'],
                          switch_notify1_email=SCRIPT_CONFIG['switch_notify1_email'],
                          switch_notify2_name=SCRIPT_CONFIG['switch_notify2_name'],
                          switch_notify2_email=SCRIPT_CONFIG['switch_notify2_email'],
                          excel_maintainer_name=SCRIPT_CONFIG['excel_maintainer_name'],
                          excel_maintainer_email=SCRIPT_CONFIG['excel_maintainer_email'],
                          script_maintainer_name=SCRIPT_CONFIG['script_maintainer_name'],
                          script_maintainer_email=SCRIPT_CONFIG['script_maintainer_email'],
                          )
    return email_body


def create_email_subject(net_vars):
    """Completes the email subject template and returns it
    
    Args: None

    Returns: STRING email_subject
    
    """
    email_subject = EMAIL_SUBJECT_TEMPLATE.format(net_vars['net_type'], net_vars['net_date'])
    logging.info("Email Subject: %s", email_subject)

    return email_subject

def gather_email_addresses():
    """Generates the email address list from the roster sheets defined
    
    Args: None
    
    Returns: STRING email_dist: Comma-separated list of emails
    
    """
    #
    # Gathering the email addresses from the Amateur Radio Roster
    #
    df_list = pd.read_excel(SCRIPT_CONFIG['roster_excel_file'],
                            sheet_name=SCRIPT_CONFIG['roster_sheet_name'],
                            skiprows=0, header=0).dropna(how='any',
                            subset=['Email']
                            )
    df_emeritus_list = pd.read_excel(SCRIPT_CONFIG['roster_excel_file'],
                                     sheet_name=SCRIPT_CONFIG['emeritus_sheet_name'],
                                     skiprows=0, header=0).dropna(how='any', subset=['Email']
                                    )

    email_dist = list(df_list['Email']) + list(df_emeritus_list['Email'])
    logging.info("Email Distribution List: %s", ",".join(email_dist))
    return email_dist

def email_net_notice (net_vars):
    """Sends an email reminder to the membership of the upcoming net

        Args: None

        Returns: None
    """

    logo = net_vars['logo']

    email_body = fill_email_net_notice_template(net_vars)
    email_subject = create_email_subject(net_vars)
    email_dist = gather_email_addresses()

    if logo is not None:
        img_data = None
        with open(logo, 'rb') as f:
            img_data = f.read()

    msg = MIMEMultipart()
    text = MIMEText(email_body, 'html')
    msg.attach(text)

    if logo is not None:
        image = MIMEImage(img_data, name=os.path.basename(logo), maintype='image',
                        subtype='png')
        msg.attach(image)

    me = SCRIPT_CONFIG['smtp_auth_user']
    logging.info("Email From: %s", me)

    if TEST is False:
        msg['Subject'] = email_subject
        msg['From'] = SCRIPT_CONFIG['email_from']

        if EMAIL_REPLY_TO is None:
            msg['Reply-to'] = SCRIPT_CONFIG['email_reply_to']

        if TEST_EMAIL is not None:
            logging.info("Sending test email")
            msg['To'] = TEST_EMAIL
            email_dist = [TEST_EMAIL]
        else:
            msg['To'] = ", ".join(email_dist)

        logger.info("Subject: %s", msg['Subject'])
        logger.info("From: %s", msg['From'])
        logger.info("To: %s", msg['To'])
        logger.info("CC: %s", msg['Cc'])
        logger.info("BCC: %s", msg['Bcc'])

        if EMAIL_REPLY_TO is not None:
            logger.info("Reply-to: %s", msg['Reply-to'],
                        )

        # Send the message via our own SMTP server, but don't include the
        # envelope header.
        s = smtplib.SMTP_SSL(SCRIPT_CONFIG['smtp_server'],
                             port=SCRIPT_CONFIG['smtp_port'])
        s.ehlo()
        s.login(me, SCRIPT_CONFIG['smtp_auth_pass'])
        s.sendmail(me, email_dist, msg.as_string())
        s.close()
    else:
        print(email_body)

def log_setup(filename):
    """Setups up timed rotation of logging
    
    Args: STRING filename Ex: test.log

    Returns: OBJECT logger
    """
    when_rotate = 'W6'
    backup_count = 2
    when_interval = 4

    log_handler =  handlers.TimedRotatingFileHandler(filename,
                                                     when=when_rotate,
                                                     interval=when_interval,
                                                     backupCount=backup_count)
    formatter = logging.Formatter(
        '%(asctime)s net_reminder.py [%(lineno)d]: %(message)s',
        '%m/%d/%Y %H:%M:%S')
    log_handler.setFormatter(formatter)
    timed_logger = logging.getLogger()
    timed_logger.addHandler(log_handler)
    timed_logger.setLevel(logging.INFO)

    return timed_logger

###########################################################
# Begin Script
###########################################################

s_net_vars = {}

# Grab the command line args
try:
    OPTS, args = getopt.getopt(sys.argv[1:], "c:l:hn:ts:e:q:x:",
                            ["help", "config=", "econfig=", "nconfig=", "log=",
                                "now=","subject=","test","test_email="]
                            )
except getopt.GetoptError as e:
    print(e)
    usage()
    sys.exit()

try:
    for o, a in OPTS:
        if o in ["-c", "--config"]:
            CONFIG_FILE = a
        elif o in ["-h", "--help"]:
            usage()
            sys.exit()
        elif o in ["-l","--log"]:
            LOG_NAME = a
        elif o in ["-e","--econfig"]:
            EMAIL_CONFIG = a
        elif o in ["-x","--nconfig"]:
            NO_NET_CONTROL_EMAIL_CONFIG = a
        elif o in ["-n","--now"]:
            NOW = datetime.strptime(a, '%m/%d/%Y')
        elif o in ["-s","--subject"]:
            EMAIL_SUBJECT_TEMPLATE = a
        elif o in ["-t","--test"]:
            TEST = True
        elif o in ["-q","--test_email"]:
            TEST_EMAIL = a
        else:
            usage()
            sys.exit()
except ValueError as e:
    print(e)
    sys.exit()

if LOG_NAME is None:
    LOG_NAME = "net_reminder.log"
if CONFIG_FILE is None:
    CONFIG_FILE = "net_reminder.yaml"
if NOW is None:
    NOW = datetime.now()

# Setup logging
logger = log_setup(LOG_NAME)
logger.info("Started")

with open(CONFIG_FILE, 'r', encoding='UTF-8') as yconfig_file:
    try:
        SCRIPT_CONFIG = yaml.safe_load(yconfig_file)

        if 'email_config' in SCRIPT_CONFIG.keys():
            EMAIL_CONFIG = SCRIPT_CONFIG['email_config']

        if 'no_net_control_email_config' in SCRIPT_CONFIG.keys():
            NO_NET_CONTROL_EMAIL_CONFIG = SCRIPT_CONFIG['no_net_control_email_config']

        if 'logo' not in SCRIPT_CONFIG.keys():
            SCRIPT_CONFIG['logo'] = None

        # If the email subject template is defined in the configuration file,
        # then use it.
        if 'email_subject_template' in SCRIPT_CONFIG.keys():
            EMAIL_SUBJECT_TEMPLATE = SCRIPT_CONFIG['email_subject_template']

        if 'no_net_control_email_subject_template' in SCRIPT_CONFIG.keys():
            NO_NET_CONTROL_EMAIL_SUBJECT_TEMPLATE = \
                SCRIPT_CONFIG['no_net_control_email_subject_template']

    except yaml.YAMLError as exc:
        print(exc)
        logger.fatal(exc)

try:
  # Use the email config default if all else fails
    if EMAIL_CONFIG is None:
        EMAIL_CONFIG = "net_reminder.html"

  # Use the email config default if all else fails
    if NO_NET_CONTROL_EMAIL_CONFIG is None:
        NO_NET_CONTROL_EMAIL_CONFIG = "no_net_reminder.html"

    # Use the email subject default if all else fails
    if EMAIL_SUBJECT_TEMPLATE is None:
        EMAIL_SUBJECT_TEMPLATE = "{0} Net for {1}"

    # Use the email subject default if all else fails
    if NO_NET_CONTROL_EMAIL_SUBJECT_TEMPLATE is None:
        NO_NET_CONTROL_EMAIL_SUBJECT_TEMPLATE = "Net for {0}"

    future_date_1wk = NOW + timedelta(days=7)
    future_date_2wk = future_date_1wk + timedelta(days=7)
    logging.info("Current Date: %s, Future Date 1wk: %s, Future Date 2wk: %s",
                 NOW, future_date_1wk, future_date_2wk)

    # Opening the Net Control Schedule to locate the Primary and Backup Net Control assignments
    df_sched = pd.read_excel(SCRIPT_CONFIG['schedule_excel_file'],
                             sheet_name=SCRIPT_CONFIG['schedule_sheet_name'],
                             skiprows=1, header=0)

    df_sched.DATE = pd.to_datetime(df_sched.DATE,format='%Y-%m-%d')

    # Get this week's Net date
    df_select_cur = df_sched[(df_sched.DATE >=
                              NOW.strptime(str(NOW)[:10], '%Y-%m-%d')) &
                             (df_sched.DATE <=
                              future_date_1wk)]

    if df_select_cur.empty:
        logging.fatal("No configured net control or backup net control for this week!",
                      )
        email_no_net_notice()
        sys.exit()

    logging.info(df_select_cur.head())

    # Get next future Net date
    df_select_next = df_sched[(df_sched.DATE >=
                               future_date_1wk.strptime(str(future_date_1wk)[:10], '%Y-%m-%d')) &
                              (df_sched.DATE <=
                               future_date_2wk)]
    if df_select_next.empty:
        logging.fatal("No configured net control or backup net control for next week!")
        email_no_net_notice()
        sys.exit()

    logging.info(df_select_next.head())

    F = "%Y-%m-%d"
    date_val_cur = datetime.strptime( str(df_select_cur['DATE'].values[0])[:10], F)
    date_val_next = datetime.strptime( str(df_select_next['DATE'].values[0])[:10], F)

    # Dates for this week and next week's Nets
    s_net_vars['net_date'] = date_val_cur.strftime("%m/%d/%Y")
    s_net_vars['next_net_date'] = date_val_next.strftime("%m/%d/%Y")

    # Current week net details
    s_net_vars['current_primary'] = df_select_cur['PRIMARY'].values[0]
    s_net_vars['current_backup'] = df_select_cur['BACKUP'].values[0]

    # Next week net details
    s_net_vars['next_primary'] = df_select_next['PRIMARY'].values[0]
    s_net_vars['next_backup'] = df_select_next['BACKUP'].values[0]

    logger.info("Date Value Current: %s, Date Value Next: %s",
                s_net_vars['net_date'],
                s_net_vars['next_net_date']
                )
    logger.info("Net Date: %s, Primary: %s, Backup: %s",
                s_net_vars['net_date'],
                s_net_vars['current_primary'],
                s_net_vars['current_backup']
                )
    logger.info("Net Date: %s, Primary: %s, Backup: %s",
                s_net_vars['next_net_date'],
                s_net_vars['next_primary'],
                s_net_vars['next_backup']
                )

    s_net_vars['net_type'] = df_select_cur['Net'].values[0]
    s_net_vars['logo'] = SCRIPT_CONFIG['logo']

    email_net_notice(s_net_vars)
except DataError as e:
    print(e)
    logging.fatal(e)
    sys.exit(e)

logging.info("Finished")
sys.exit()
