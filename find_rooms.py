#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import subprocess
import getpass
import xml.etree.ElementTree as ET
import argparse
import csv
import operator
from string import Template
import string

def findRooms(prefix):
    rooms = {}
    data = xml.substitute(name=prefix)

    header = "\"content-type: text/xml;charset=utf-8\""
    command = f"curl --silent --header {header} --data '{data}' --ntlm --negotiate -u '{user}:{password}' {url}"
    response = subprocess.Popen(command, stdout=subprocess.PIPE, shell=True).communicate()[0]
    print(f"XML return: {data}")
    tree = ET.fromstring(response)

    elems = tree.findall(".//{http://schemas.microsoft.com/exchange/services/2006/types}Resolution")
    for elem in elems:
        email = elem.findall(".//{http://schemas.microsoft.com/exchange/services/2006/types}EmailAddress")
        name = elem.findall(".//{http://schemas.microsoft.com/exchange/services/2006/types}DisplayName")
        if len(email) > 0 and len(name) > 0:
            rooms[email[0].text] = name[0].text
    return rooms

parser = argparse.ArgumentParser()
parser.add_argument("prefix", nargs='+', help="A list of prefixes to search for. E.g. 'conference confi'")
parser.add_argument("-url", "--url", help="URL for Exchange server, e.g. 'https://mail.domain.com/ews/exchange.asmx'.", required=True)
parser.add_argument("-u", "--user", help="User name for Exchange/Outlook", required=True)
parser.add_argument("-d", "--deep", help="Attempt a deep search (takes longer).", action="store_true")
args = parser.parse_args()

url = args.url
user = args.user
password = getpass.getpass("Password:")

with open("resolvenames_template.xml", "r", encoding='utf-8') as template_file:
    xml_template = template_file.read()
xml = Template(xml_template)

rooms = {}

for prefix in args.prefix:
    rooms.update(findRooms(prefix))
    print(f"After searching for prefix '{prefix}' we found {len(rooms)} rooms.")

    if args.deep: 
        symbols = string.ascii_letters + string.digits
        for symbol in symbols:
            prefix_deep = f"{prefix} {symbol}"
            rooms.update(findRooms(prefix_deep))

        print(f"After deep search for prefix '{prefix}' we found {len(rooms)} rooms.")

# Writing to CSV file
with open("rooms.csv", "w", newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)
    for item in sorted(rooms.items(), key=operator.itemgetter(1)):
        writer.writerow([item[1], item[0]])
