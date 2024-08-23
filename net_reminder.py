"""

Script: net_reminder.py
Purpose: To send an HTML email weekly based on a user-defined template

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
import os
import sys
# Import smtplib for the actual sending function
import smtplib
# Email templating
from jinja2 import Template
# Easy handling of data
import pandas as pd
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
NOW = None
TEST = False
EMAIL_SUBJECT_TEMPLATE = None
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

###########################################################
# Begin Script
###########################################################

# Grab the command line args
try:
    OPTS, args = getopt.getopt(sys.argv[1:], "c:l:hn:ts:e:q:",
                               ["help", "config=", "econfig=", "log=",
                                "now=","subject=","test","test_email="]
                               )

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
        elif o in ["-n","--now"]:
            NOW = datetime.strptime(a, '%m/%d/%Y')
        elif o in ["-s","--subject"]:
            EMAIL_SUBJECT_TEMPLATE = a
        elif o in ["-t","--test"]:
            TEST = True
        elif o in ["-q","--test_email"]:
            TEST_EMAIL = a
        else:
            print("Unhandled option {o} on command line")
            assert False, "Unhandled option {o} on command line"

except Exception as e:
    print(e)
    sys.exit(1)

if LOG_NAME is None:
    LOG_NAME = "net_reminder.log"
if CONFIG_FILE is None:
    CONFIG_FILE = "net_reminder.yaml"
if NOW is None:
    NOW = datetime.now()

logger = logging.getLogger(__name__)
user = os.environ['USER']
FORMAT = '%(asctime)s %(user)-8s %(message)s'
d = {'user': user}
logging.basicConfig(filename=LOG_NAME, format=FORMAT, level=logging.INFO)
logger.info("Started", extra=d)

with open(CONFIG_FILE, 'r', encoding='UTF-8') as yconfig_file:
    try:
        SCRIPT_CONFIG = yaml.safe_load(yconfig_file)

        if 'EMAIL_CONFIG' in SCRIPT_CONFIG.keys():
            EMAIL_CONFIG = SCRIPT_CONFIG['email_config']

        # If the email subject template is defined in the configuration file,
        # then use it.
        if 'email_subject_template' in SCRIPT_CONFIG.keys():
            EMAIL_SUBJECT_TEMPLATE = SCRIPT_CONFIG['email_subject_template']

    except yaml.YAMLError as exc:
        print(exc)
        logger.fatal(exc, extra=d)

try:
  # Use the email config default if all else fails
    if EMAIL_CONFIG is None:
        EMAIL_CONFIG = "net_reminder.html"

    # Use the email subject default if all else fails
    if EMAIL_SUBJECT_TEMPLATE is None:
        EMAIL_SUBJECT_TEMPLATE = "{0} Net for {1}"

    logo = SCRIPT_CONFIG['logo']

    future_date_1wk = NOW + timedelta(days=7)
    future_date_2wk = future_date_1wk + timedelta(days=7)
    logging.info("Current Date: %s, Future Date 1wk: %s, Future Date 2wk: %s",
                 NOW, future_date_1wk, future_date_2wk, extra=d)

    # Opening the Net Control Schedule to locate the Primary and Backup Net Control assignments
    df_sched = pd.read_excel(SCRIPT_CONFIG['schedule_excel_file'],
                             sheet_name=SCRIPT_CONFIG['schedule_sheet_name'],
                             skiprows=1, header=0)

    df_sched.DATE = pd.to_datetime(df_sched.DATE,format='%Y-%m-%d')

    df_select_cur = df_sched[(df_sched.DATE >= NOW.strptime(str(NOW)[:10], '%Y-%m-%d')) &
                             (df_sched.DATE <= future_date_1wk)]
    logging.info(df_select_cur.head(), extra=d)

    df_select_next = df_sched[(df_sched.DATE >= future_date_1wk.strptime(str(future_date_1wk)[:10], '%Y-%m-%d')) &
                              (df_sched.DATE <= future_date_2wk)]
    logging.info(df_select_next.head(), extra=d)

    F = "%Y-%m-%d"
    date_val_cur = datetime.strptime( str(df_select_cur['DATE'].values[0])[:10], F)
    date_val_next = datetime.strptime( str(df_select_next['DATE'].values[0])[:10], F)

    current_net_date = date_val_cur.strftime("%m/%d/%Y")
    current_primary = df_select_cur['PRIMARY'].values[0]
    current_backup = df_select_cur['BACKUP'].values[0]

    next_net_date = date_val_next.strftime("%m/%d/%Y")
    next_primary = df_select_next['PRIMARY'].values[0]
    next_backup = df_select_next['BACKUP'].values[0]

    logger.info("Date Value Current: %s, Date Value Next: %s", date_val_cur, date_val_next,
                extra=d)
    logger.info("Net Date: %s, Primary: %s, Backup: %s",
                current_net_date,
                current_primary,
                current_backup,
                extra=d)
    logger.info("Net Date: %s, Primary: %s, Backup: %s",
                next_net_date,
                next_primary,
                next_backup,
                extra=d)

    with open(EMAIL_CONFIG, 'r', encoding='UTF-8') as email_template:
        try:
            TEMPLATE = email_template.read()
            email_template.close()
        except Exception as exc:
            print(exc)
            logger.warning(exc)
            TEMPLATE = DEFAULT_EMAIL_TEMPLATE

    #
    # HTML email template
    #
    t = Template(TEMPLATE)
    email_body = t.render(net_date=date_val_cur.strftime("%m/%d/%Y"),
                          primary_net_control=df_select_cur['PRIMARY'].values[0],
                          backup_net_control=df_select_cur['BACKUP'].values[0],
                          net_type=df_select_cur['Net'].values[0],logo=os.path.basename(logo),
                          primary_net_control_2wk=df_select_next['PRIMARY'].values[0],
                          backup_net_control_2wk=df_select_next['BACKUP'].values[0],
                          net_date_2wk=date_val_next.strftime("%m/%d/%Y"),
                          switch_notify1_name=SCRIPT_CONFIG['switch_notify1_name'],
                          switch_notify1_email=SCRIPT_CONFIG['switch_notify1_email'],
                          switch_notify2_name=SCRIPT_CONFIG['switch_notify2_name'],
                          switch_notify2_email=SCRIPT_CONFIG['switch_notify2_email'],
                          excel_maintainer_name=SCRIPT_CONFIG['excel_maintainer_name'],
                          excel_maintainer_email=SCRIPT_CONFIG['excel_maintainer_email'],
                          script_maintainer_name=SCRIPT_CONFIG['script_maintainer_name'],
                          script_maintainer_email=SCRIPT_CONFIG['script_maintainer_email'],
                          )
    net_type = df_select_cur['Net'].values[0]
    net_date = date_val_cur.strftime("%m/%d/%Y")
    email_subject = EMAIL_SUBJECT_TEMPLATE.format(net_type, net_date)
    logging.info("Email Subject: %s", email_subject, extra=d)

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
    logging.info("Email Distribution List: %s", ",".join(email_dist),
                 extra=d)

    IMG_DATA = None
    with open(logo, 'rb') as f:
        IMG_DATA = f.read()

    msg = MIMEMultipart()
    text = MIMEText(email_body, 'html')
    msg.attach(text)
    image = MIMEImage(IMG_DATA, name=os.path.basename(logo), maintype='image',
                      subtype='png')
    msg.attach(image)

    me = SCRIPT_CONFIG['smtp_auth_user']
    logging.info("Email From: %s", me ,extra=d)

    if TEST is False:
        msg['Subject'] = email_subject
        msg['From'] = SCRIPT_CONFIG['email_from']

        if EMAIL_REPLY_TO is None:
            msg['Reply-to'] = SCRIPT_CONFIG['email_reply_to']

        if TEST_EMAIL is not None:
            msg['To'] = TEST_EMAIL
            email_dist = [TEST_EMAIL]
        else:
            msg['To'] = ", ".join(email_dist)

        logger.info("Subject: %s", msg['Subject'], extra=d)
        logger.info("From: %s", msg['From'], extra=d)
        logger.info("To: %s", msg['To'], extra=d)
        logger.info("CC: %s", msg['Cc'], extra=d)
        logger.info("BCC: %s", msg['Bcc'], extra=d)

        if EMAIL_REPLY_TO is not None:
            logger.info("Reply-to: %s", msg['Reply-to'],
                        extra=d)

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
except Exception as e:
    print(e)
    logging.fatal(e, extra=d)
    sys.exit(e)

logging.info("Finished", extra=d)
sys.exit()
