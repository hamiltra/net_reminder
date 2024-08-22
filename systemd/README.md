# TeamTalk 5 systemd service for Debian 9

Here's the instruction for installing the Net Reminder script on Ubuntu 22.04 LTS.

## Configure the Net Reminder service

Edit `net_reminder.service` and set up the paths to binary,
yaml configuration and log file. Basically the line starting with
`ExecStart=`.

Edit `net_reminder.timer` and set the OnCalendar spec to your need(s).

## Installing and starting the Net Reminder service

1. Install the net_reminder.py and net_reminder.sh scripts into the /usr/local/bin directory
2. Set permissions to 0755 on the scripts
3. Add netreminder user and group: useradd -m /home/net_reminder -rs /bin/bash netreminder
4. In the netreminder user directory, create the directories: excel_src, html_src and image_src.
5. Copy the Excel, HTML email templates and image files to their respective directories
6. Set any needed permissions on the files and folders
7. Create the directory /etc/net_reminder and copy the net_reminder.yaml file into it
8. Again, set the permissions to be restrictive such as chmod 0640 on the configuration file
9. Create the log directory /var/log/net_reminder
10. touch /var/log/net_reminder.log
11. Set permissions and ownership on the log for netreminder

Afterwards as root user, copy `net_reminder.service` to `/etc/systemd/system` and set permissions.

As root user enable the Net Reminder service:

`systemctl enable net_reminder`
`systemctl enable net_reminder.timer`

As root user start the Net Reminder timer service:

`systemctl start net_reminder.timer`

As root user, test the Net Reminder service and verify that the email received is satisfactory

`systemctl start net_reminder`

## Stopping and uninstalling the Net Reminder service

To stop the TeamTalk service:

`systemctl stop net_reminder`

To uninstall the TeamTalk service:

`systemctl disable net_reminder`

Undo the installation steps.
