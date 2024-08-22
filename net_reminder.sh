#!/bin/bash

echo "Starting net_reminder"

# Go to home directory
pushd $HOME

# Activate virtual environment
source venv/bin/activate

# Setup virtual environment
pip install -r requirements.txt

# Run net_reminder
#
# For testing...
# python3 /usr/local/bin/net_reminder.py -c /etc/net_reminder/net_reminder.yaml --log /var/log/net_reminder/net_reminder.log --test --test_email <your email>

# For running...
python3 /usr/local/bin/net_reminder.py -c /etc/net_reminder/net_reminder.yaml --log /var/log/net_reminder/net_reminder.log 

# Deactivate the virtual environment
deactivate

exit