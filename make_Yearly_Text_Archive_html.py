#!/usr/local/bin/python
'''
Yahoo-Groups-Archiver, HTML Archive Script Copyright 2019 Robert Lancaster and others

YahooGroups-Archiver, a simple python script that allows for all
messages in a public Yahoo Group to be archived.

The HTML Archive Script allows you to take the downloaded json documents
and turn them into html-based yearly archives of emails.
Note that the archive-group.py script must be run first.

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
'''
import email
import HTMLParser
import json
import os
import sys
from datetime import datetime
from natsort import natsorted, ns
import cgi

#To avoid Unicode Issues
reload(sys)
sys.setdefaultencoding('utf-8')

def archiveYahooMessage(file, archiveDir, messageYear, format):
     try:
          archiveYear = archiveDir + '/archive-' + str(messageYear) + '.html'
          messageText = loadYahooMessage(file, format)

          f = open(archiveYear, 'a')
          if f.tell() == 0:
               f.write("<style>pre {white-space: pre-wrap;}</style>\n");
          f.write(messageText)
          f.close()
          print 'Yahoo Message: ' + file + ' archived to: archive-' + str(messageYear) + '.html'
     except Exception as e:
          print 'Yahoo Message: ' + file + ' had an error:'
          print e

def loadYahooMessage(file, format):
    f1 = open(file,'r')
    fileContents=f1.read()
    f1.close()
    jsonDoc = json.loads(fileContents)
    
    if 'ygData' not in jsonDoc:
         return ''
    
    emailMessageID = jsonDoc['ygData']['msgId']
    emailMessageSender = HTMLParser.HTMLParser().unescape(jsonDoc['ygData']['from']).decode(format).encode('utf-8')
    emailMessageTimeStamp = jsonDoc['ygData']['postDate']
    emailMessageDateTime = datetime.fromtimestamp(float(emailMessageTimeStamp)).strftime('%Y-%m-%d %H:%M:%S')
    emailMessageSubject = HTMLParser.HTMLParser().unescape(jsonDoc['ygData']['subject']).decode(format).encode('utf-8')
    emailMessageString = HTMLParser.HTMLParser().unescape(jsonDoc['ygData']['rawEmail']).decode(format).encode('utf-8')
    message = email.message_from_string(emailMessageString)
    messageBody = getEmailBody(message)

    messageText =  ''
    messageText += '<font color="#0033cc">\n'
    messageText += '-----------------------------------------------------------------------------------<br>' + "\n"
    messageText += 'Post ID: ' + str(emailMessageID) + '<a name=\"' + str(emailMessageID) + '\"></a><br>' + "\n"
    messageText += 'Sender: ' + cgi.escape(emailMessageSender) + '<br>' + "\n"
    messageText += 'At: ' + cgi.escape(emailMessageDateTime) + '<br>' + "\n"
    messageText += 'Subject: ' + cgi.escape(emailMessageSubject) + '<br>' + "\n"
    messageText += '<br>' + "\n"
    messageText += '</font>\n'
    messageText += messageBody
    messageText += '<br><br><br><br><br>' + "\n"
    return messageText
    
def getYahooMessageYear(file):
    f1 = open(file,'r')
    fileContents=f1.read()
    f1.close()
    jsonDoc = json.loads(fileContents)
    if 'ygData' not in jsonDoc:
         return None
    emailMessageTimeStamp = jsonDoc['ygData']['postDate']
    return datetime.fromtimestamp(float(emailMessageTimeStamp)).year

# Thank you to the help in this forum for the bulk of this function
# https://stackoverflow.com/questions/17874360/python-how-to-parse-the-body-from-a-raw-email-given-that-raw-email-does-not
def getEmailBody(message):
    body = ''
    if message.is_multipart():
        for part in message.walk():
            ctype = part.get_content_type()
            cdispo = str(part.get('Content-Disposition'))

            # skip any text/plain (txt) attachments
            if ctype == 'text/plain' and 'attachment' not in cdispo:
                body += '<pre>'
                body += cgi.escape(part.get_payload(decode=True))  # decode
                body += '</pre>'
                break
    # not multipart - i.e. plain text, no attachments, keeping fingers crossed
    else:
        ctype = message.get_content_type()
        if ctype != 'text/html':
             body += '<pre>'
             body += cgi.escape(message.get_payload(decode=True))
             body += '</pre>'
        else:
             body += message.get_payload(decode=True)
    return body

## This is where the script starts

if len(sys.argv) < 2:
     sys.exit('You need to specify your group name')

groupName = sys.argv[1]
oldDir = os.getcwd()
if os.path.exists(groupName):
    archiveDir = os.path.abspath(groupName + '-archive')
    if not os.path.exists(archiveDir):
         os.makedirs(archiveDir)
    os.chdir(groupName)
    for file in natsorted(os.listdir(os.getcwd())):
         messageYear = getYahooMessageYear(file)
         if messageYear:
              archiveYahooMessage(file, archiveDir, messageYear, 'utf-8')
else:
     sys.exit('Please run archive-group.py first')

os.chdir(oldDir)
print('Complete')


