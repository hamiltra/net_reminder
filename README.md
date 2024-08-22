# Project: net_reminder

### Net Reminder
The purpose of this script is to provide an HTML email once a week as a reminder to staff the weekly nets for the organization. To accomplish this, there are two Excel spreadsheets that are required. One file, the Net Schedule is in one Excel file. The Roster containing the list of members is in another file. The script operates by looking up the weekly net based on current time, crafting an HTML email based from the configured template, and distributing to an email list created from 2 roster sheets.

### Installing
Installation tasks
1. Download the release distribution or clone the repository
2. Create a login account if necessary.
3. Setup the configuration options including the file paths
4. Set permissions on the configuration options file for the user. Turn off world readble/writable. For example, chnmod 0750
5. Install and configure the system timer scripts
6. Trigger the script and verify logs
7. Test and verify that the information is how you desire it to be while making any adjustments

### Running
Below is a snapshot of the script help.

Script Help
```
net_reminder.py - Net Reminder

 -h, --help        This notice.
 -c, --config      Configuration file. default: net_reminder.yaml
 -e, --econfig     Email layout configuration file. default: net_reminder.html
 -l, --log         Log file. default: net_reminder.log
 -n, --now         Now: mm/dd/yyyy. default: current date
 -s, --subject     Email subject. default: '<0> Net for <1>'
 -t, --test        Run script but don't send mail and print output. default: False
 -q, --test_email  Run script using test emails provided. default: None
```

Here are a few examples of how to run and to test your configurations before going to production.
```
python3 net_reminder.py -c net_reminder_org.yaml --test_email xyzzy@example.com --test
python3 net_reminder.py -c net_reminder_org.yaml --test_email xyzzy@example.com
python3 net_reminder.py -c net_reminder_org.yaml -n 09/18/2024 --test_email xyzzy@example.com
python3 net_reminder.py -c net_reminder_org.yaml --log test.log --test_email xyzzy@example.com
```

#### Email Template
The email template is configurable as an HTML template. The default file is net_reminder.html. As such, there are several variables that are available to the template. Static variables are managed in the configuration file. Dynamic variables are determined at run-time.

Template Variables:
* net_type (dynamic)
* net_date (dynamic)
* primary_net_control (dynamic)
* backup_net_control (dynamic)
* primary_net_control_2wk (dynamic)
* backup_net_control_2wk (dynamic)
* switch_notify1_name
* switch_notify1_email
* switch_notify2_name
* switch_notify2_email
* excel_maintainer_name
* excel_maintainer_email
* script_maintainer_name
* script_maintainer_email

Sample Email Template
```<!DOCTYPE html>
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
  Weekly Nets and Travel Nets are held on repeater: <b><font color="red">XXX.XXX- PL NNN.N</font></b>. After the Net or any activity, please login to the website to record your time in Time Manager.
  <br><br>
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
<font align='center' class="font-small">This notice automated using the Amateur Radio net_reminder script <br>maintained by {{script_maintainer_name}}, <a href='mailto: {{script_maintainer_email}}'>{{script_maintainer_email}}</a></font><br>
</td></tr>
</table>
{% endblock %}
</body>
</html>
```

### Configuration File
The default configuration file is net_reminder.yaml. All of the configurable options are set within this file.

Sample Configuration Template:
```#######################################
#
# Excel Spreadsheet Configuration
#
#######################################
# Excel Roster and Schedule files and sheets
schedule_excel_file: excel_src/NetControlSchedule.xlsx
schedule_sheet_name: Rev 2
roster_excel_file: excel_src/Roster.xlsx
roster_sheet_name: Active
emeritus_sheet_name: Emeritus
#######################################
#
# Email Configuration
#
#######################################
# Email SMTP configuration setup
email_from: <Email from>
smtp_server: <SMTP Server>
smtp_port: <SMTP Server Port>
smtp_auth_user: <Authentication user>
smtp_auth_pass: <Authentication password>
# Email logo file
logo: image_src/LNLogo.png
# Email subject template
email_subject_template: {0} Net for {1}
# Email body template file
email_config: html_src/net_reminder.html
# Email reply-to
email_reply_to: <Who should receive reply emails>
# Email body script maintainer info
script_maintainer_name: Your Name
script_maintainer_email: xyzzy@example.com
# Email body Excel files maintainer info
excel_maintainer_name: <Excel file maintainer>
excel_maintainer_email: <Excel file email address>
# Email body switching duty notification maintainer
switch_notify1_name: <Name of who to notify when switching>
switch_notify1_email: <Email of who to notify when switching>
switch_notify2_name: Your Name
switch_notify2_email: xyzzy@example.com
```