###############################################################################
#
# Script: net_reminder.py
# Purpose: To send an HTML email weekly based on a user-defined template
#
###############################################################################
# Date handling
from datetime import datetime, timedelta
# Import the email modules we'll need
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
# Command line args
import getopt
# Email templating
from jinja2 import Template
import logging
import os
# Easy handling of data
import pandas as pd
# Import smtplib for the actual sending function
import smtplib
import sys
# Configuration file
import yaml

#######################################
# Declaring global vars
#######################################
script_config = None
opts = None
ret_val = -1

#######################################
# Command Line Options
#######################################
log_name = None
config_file = None
email_config = None
now = None
test = False
email_subject_template = None
test_email = None
email_reply_to = None

#######################################
# Sample email template
#######################################
default_email_template = '''\
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
  opts, args = getopt.getopt(sys.argv[1:], "c:l:hn:ts:e:q:", ["help", "config=", "econfig=", "log=", "now=","subject=","test","test_email="])
  
  for o, a in opts:
    if o in ["-c", "--config"]:
      config_file = a
    elif o in ["-h", "--help"]:
      usage()
      sys.exit()
    elif o in ["-l","--log"]:
      log_name = a
    elif o in ["-e","--econfig"]:
      email_config = a
    elif o in ["-n","--now"]:
      now = datetime.strptime(a, '%m/%d/%Y')
    elif o in ["-s","--subject"]:
      email_subject_template = a
    elif o in ["-t","--test"]:
      test = True
    elif o in ["-q","--test_email"]:
      test_email = a
    else:
      print("Unhandled option %s on command line" % o)
      assert False, "Unhandled option %s on command line" % o

except Exception as e:
  print(e)
  sys.exit(1)

if log_name == None:
  log_name = "net_reminder.log"
if config_file == None:
  config_file = "net_reminder.yaml"
if now == None:
  now = datetime.now()

logger = logging.getLogger(__name__)
user = os.environ['USER']
FORMAT = '%(asctime)s %(user)-8s %(message)s'
d = {'user': user}
logging.basicConfig(filename=log_name, format=FORMAT, level=logging.INFO)
logger.info("Started", extra=d)

with open(config_file, 'r') as config_file:
  try: 
    script_config = yaml.safe_load(config_file)
    
    if 'email_config' in script_config.keys():
      email_config = script_config['email_config']

    # If the email subject template is defined in the configuration file,
    # then use it.
    if 'email_subject_template' in script_config.keys():
      email_subject_template = script_config['email_subject_template']

  except yaml.YAMLError as exc:
    print(exc)
    logger.fatal(exc, extra=d)

try: 
  # Use the email config default if all else fails
  if email_config == None:
    email_config = "net_reminder.html"

  # Use the email subject default if all else fails
  if email_subject_template == None:
    email_subject_template = "{0} Net for {1}"
      
  logo = script_config['logo']

  future_date_1wk = now + timedelta(days=7)
  future_date_2wk = future_date_1wk + timedelta(days=7)
  logging.info("Current Date: %s, Future Date 1wk: %s, Future Date 2wk: %s" % (now, future_date_1wk, future_date_2wk), extra=d)

  # Opening the Net Control Schedule to locate the Primary and Backup Net Control assignments
  df_sched = pd.read_excel(script_config['schedule_excel_file'], sheet_name=script_config['schedule_sheet_name'], skiprows=1, header=0)

  df_sched.DATE = pd.to_datetime(df_sched.DATE,format='%Y-%m-%d')

  df_select_cur = df_sched[(df_sched.DATE >= now.strptime(str(now)[:10], '%Y-%m-%d')) & (df_sched.DATE <= future_date_1wk)]
  logging.info(df_select_cur.head(), extra=d)

  df_select_next = df_sched[(df_sched.DATE >= future_date_1wk.strptime(str(future_date_1wk)[:10], '%Y-%m-%d')) & (df_sched.DATE <= future_date_2wk)]
  logging.info(df_select_next.head(), extra=d)

  f = "%Y-%m-%d"
  date_val_cur = datetime.strptime( str(df_select_cur['DATE'].values[0])[:10], f)
  date_val_next = datetime.strptime( str(df_select_next['DATE'].values[0])[:10], f)
  
  logger.info("Date Value Current: %s, Date Value Next: %s" % ( str(date_val_cur), str(date_val_next)), extra=d)
  logger.info("Net Date: %s, Primary: %s, Backup: %s" % ( date_val_cur.strftime("%m/%d/%Y"), df_select_cur['PRIMARY'].values[0], df_select_cur['BACKUP'].values[0]), extra=d)
  logger.info("Net Date: %s, Primary: %s, Backup: %s" % ( date_val_next.strftime("%m/%d/%Y"), df_select_next['PRIMARY'].values[0], df_select_next['BACKUP'].values[0]), extra=d)

  with open(email_config, 'r') as email_template:
    try: 
      template = email_template.read()
      email_template.close()
    except Exception as exc:
      print(exc)
      logger.warn(exc)
      template = default_email_template

  #
  # HTML email template
  #
  t = Template(template)
  email_body = t.render(net_date=date_val_cur.strftime("%m/%d/%Y"), primary_net_control=df_select_cur['PRIMARY'].values[0], 
                backup_net_control=df_select_cur['BACKUP'].values[0], net_type=df_select_cur['Net'].values[0],logo=os.path.basename(logo),
                primary_net_control_2wk=df_select_next['PRIMARY'].values[0], backup_net_control_2wk=df_select_next['BACKUP'].values[0], net_date_2wk=date_val_next.strftime("%m/%d/%Y"),
                switch_notify1_name=script_config['switch_notify1_name'], switch_notify1_email=script_config['switch_notify1_email'],
                switch_notify2_name=script_config['switch_notify2_name'], switch_notify2_email=script_config['switch_notify2_email'],
                excel_maintainer_name=script_config['excel_maintainer_name'], excel_maintainer_email=script_config['excel_maintainer_email'],
                script_maintainer_name=script_config['script_maintainer_name'], script_maintainer_email=script_config['script_maintainer_email'],
  )
  net_type = df_select_cur['Net'].values[0]
  net_date = date_val_cur.strftime("%m/%d/%Y")
  email_subject = email_subject_template.format(net_type, net_date)
  logging.info("Email Subject: " + email_subject, extra=d)

  #
  # Gathering the email addresses from the Amateur Radio Roster
  #
  df_list = pd.read_excel(script_config['roster_excel_file'], sheet_name=script_config['roster_sheet_name'], skiprows=0, header=0).dropna(how='any', subset=['Email'])
  df_emeritus_list = pd.read_excel(script_config['roster_excel_file'], sheet_name=script_config['emeritus_sheet_name'], skiprows=0, header=0).dropna(how='any', subset=['Email'])

  email_dist = list(df_list['Email']) + list(df_emeritus_list['Email'])
  logging.info("Email Distribution List: %s" % ",".join(email_dist), extra=d)

  img_data = None
  with open(logo, 'rb') as f:
      img_data = f.read()

  msg = MIMEMultipart()
  text = MIMEText(email_body, 'html')
  msg.attach(text)
  image = MIMEImage(img_data, name=os.path.basename(logo), maintype='image', subtype=
                    'png')
  msg.attach(image)

  me = script_config['smtp_auth_user']
  logging.info("Email From: " + me, extra=d)

  if test == False:
    msg['Subject'] = email_subject
    msg['From'] = script_config['email_from']
    
    if email_reply_to == None:
      msg['Reply-to'] = script_config['email_reply_to']
    
    if test_email != None:
      msg['To'] = test_email
      email_dist = [test_email]
    else:
      msg['To'] = ", ".join(email_dist)

    logger.info("Subject: %s" % (msg['Subject']), extra=d)
    logger.info("From: %s" % (msg['From']), extra=d)
    logger.info("To: %s" % (msg['To']), extra=d)
    logger.info("CC: %s" % (msg['Cc']), extra=d)
    logger.info("BCC: %s" % (msg['Bcc']), extra=d)

    if email_reply_to != None:
      logger.info("Reply-to: %s" % (msg['Reply-to']), extra=d)

    # Send the message via our own SMTP server, but don't include the
    # envelope header.
    s = smtplib.SMTP_SSL(script_config['smtp_server'], port=script_config['smtp_port'])
    s.ehlo()
    s.login(me, script_config['smtp_auth_pass'])
    s.sendmail(me, email_dist, msg.as_string())
    s.close()
  else:
    print(email_body)
  ret_val = 0
except Exception as e:
  print(e)
  logging.fatal(e, extra=d)
  ret_val = 1

logging.info("Finished", extra=d)
sys.exit(ret_val)