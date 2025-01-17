#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import subprocess
import getpass
from string import Template
import xml.etree.ElementTree as ET
import csv
import argparse
import datetime

now = datetime.datetime.now().replace(microsecond=0)
starttime_default = now.isoformat()
end_time_default = None

parser = argparse.ArgumentParser()
parser.add_argument("-url", "--url", help="URL for Exchange server, e.g. 'https://mail.domain.com/ews/exchange.asmx'.", required=True)
parser.add_argument("-u", "--user", help="User name for Exchange/Outlook", required=True)
parser.add_argument("-start", "--starttime", help="Starttime e.g. 2014-07-02T11:00:00 (default = now)", default=starttime_default)
parser.add_argument("-end", "--endtime", help="Endtime e.g. 2014-07-02T12:00:00 (default = now+1h)", default=end_time_default)
#parser.add_argument("-n","--now", help="Will set starttime to now and endtime to now+1h", action="store_true")
parser.add_argument("-f", "--file", help="CSV filename with rooms to check (default=rooms.csv). Format: Name,email", default="rooms.csv")

args = parser.parse_args()

url = args.url

rooms = {}
with open(args.file, 'r', encoding='utf-8') as file:
    reader = csv.reader(file)
    for row in reader:
        rooms[row[1]] = row[0]

start_time = args.starttime
if not args.endtime:
    start = datetime.datetime.strptime(start_time, "%Y-%m-%dT%H:%M:%S")
    end_time = (start + datetime.timedelta(hours=1)).isoformat()
else:
    end_time = args.endtime

user = args.user
password = getpass.getpass("Password:")

print(f"Searching for a room from {start_time} to {end_time}:")
print(f"{'Status':<10} {'Room':<64} {'Email':<64}")

with open("getavailibility_template.xml", "r", encoding='utf-8') as template_file:
    xml_template = template_file.read()

xml = Template(xml_template)
for room in rooms:
    data = xml.substitute(email=room, starttime=start_time, endtime=end_time)

    header = "\"content-type: text/xml;charset=utf-8\""
    command = f"curl --silent --header {header} --data '{data}' --ntlm --negotiate -u '{user}:{password}' {url}"
    response = subprocess.Popen(command, stdout=subprocess.PIPE, shell=True).communicate()[0]

    tree = ET.fromstring(response)

    status = "Free"
    # Namespaces are tricky with XML, but this works for the Exchange format
    elems = tree.findall(".//{http://schemas.microsoft.com/exchange/services/2006/types}BusyType")
    for elem in elems:
        status = elem.text

    print(f"{status:<10} {rooms[room]:<64} {room:<64}")

