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
from collections import OrderedDict
import argparse

#To avoid Unicode Issues
reload(sys)
sys.setdefaultencoding("utf-8")

Messages = OrderedDict()
Threads = OrderedDict()
Senders = OrderedDict()

class MessageData:
     def __init__(self):
          pass

class ThreadData:
     def __init__(self):
          pass

class SenderData:
     def __init__(self):
          pass
     
def senderName(messageSender):
     # Trim past '<' character, if any
     idx = messageSender.find("<")
     if (idx >= 0):
          messageSender = messageSender[:idx]

     idx = messageSender.find("@")
     if (idx >= 0):
          messageSender = messageSender[:idx]
          
     # Trim trailing spaces, starting quotes and ending quotes
     messageSender = messageSender.rstrip()
     messageSender = messageSender.rstrip('\"')
     messageSender = messageSender.lstrip('\"')

     return messageSender

def senderIndexFileName(messageSenderName):
     return messageSenderName.replace(" ", "_")

def archiveYahooMessage(messageId, archiveDir, format):
     try:
          fileName = "{}.json".format(messageId)
          archiveMessageFile = "{}/{}.html".format(archiveDir, messageId)
          
          messageText = loadYahooMessage(fileName, format)

          if not messageText:
               print("Yahoo Message: {} skipped".format(fileName))
               return

          # Update the archive file
          f = open(archiveMessageFile, "w")
          if f.tell() == 0:
               f.write("<style>pre {white-space: pre-wrap;}</style>\n");
          f.write(messageText)
          f.close()
          
          print("Yahoo Message: {} archived to: {}.html".format(fileName, messageId))               
     except Exception as e:
          print("Yahoo Message: {} had an error: {}".format(fileName, e))

def archiveYahooByThread(year, archiveDir, format):
     try:
          archiveThreadFile = "{}/thread-index-{}.html".format(archiveDir, year)

          # Update the archive file
          f = open(archiveThreadFile, "w")
          f.write("<style>pre {white-space: pre-wrap;}</style>\n");
          f.write("<h3>{}@yahoogroups.com threads in {}</h3>\n".format(groupName, year));

          f.write("[<a href='index.html'>All Years</a>] ")
          
          for y in Threads:
               if y != year:
                    f.write("[<a href='thread-index-{}.html'>{}</a>] ".format(y, y))
               else:
                    f.write("[{}] ".format(y))
          
          f.write("<br><br>\n");
          f.write("<ul>\n");

          for threadId in Threads[year]:
               messageId = threadId
               messageSubject = Messages[messageId].messageSubject
               messageSender = Messages[messageId].messageSender
               messageSenderName = Messages[messageId].messageSenderName
               messageTimeStamp = Messages[messageId].messageTimeStamp
               #messageDateTime = datetime.fromtimestamp(float(messageTimeStamp)).strftime("%b %-d, %Y")
               f.write(" <li><a name='{}'></a><a href='{}.html'>{}</a>, <em>{}</em>\n".format(threadId, threadId, cgi.escape(messageSubject), cgi.escape(messageSenderName)));
               if Messages[threadId].messageThreadNext:
                    messageId = Messages[threadId].messageThreadNext
                    f.write(" <ul>\n");
                    while messageId:
                         messageSubject = Messages[messageId].messageSubject
                         messageSender = Messages[messageId].messageSender
                         messageTimeStamp = Messages[messageId].messageTimeStamp
                         #messageDateTime = datetime.fromtimestamp(float(messageTimeStamp)).strftime("%b %-d, %Y")
                         f.write("  <li><a name='{}'></a><a href='{}.html'>{}</a>, <em>{}</em>\n".format(messageId, messageId, cgi.escape(messageSubject), cgi.escape(messageSenderName)));
                         messageId = Messages[messageId].messageThreadNext
                    f.write(" </ul>\n");
                    
          f.write("</ul>\n");

          f.write("<br><br>Generated by the <a href='https://github.com/andrei-radulescu-banu/YahooGroups-Archiver'>YahooGroups-Archiver</a> on {}.".format(datetime.now().strftime("%b %-d, %Y")))

          f.close()

          print("Yahoo Thread archived to thread-index-{}.html".format(year))

     except Exception as e:
          print("Yahoo Message: {} had an error: {}".format(archiveThreadFile, e))
    
def archiveYahooByDate(year, archiveDir, format):
     try:
          archiveDateFile = "{}/date-index-{}.html".format(archiveDir, year)

          # Update the archive file
          f = open(archiveDateFile, "w")
          f.write("<style>pre {white-space: pre-wrap;}</style>\n");
          f.write("<h3>{}@yahoogroups.com messages by date in {}</h3>\n".format(groupName, year));

          f.write("[<a href='index.html'>All Years</a>] ")
          
          for y in Threads:
               if y != year:
                    f.write("[<a href='date-index-{}.html'>{}</a>] ".format(y, y))
               else:
                    f.write("[{}] ".format(y))
          
          f.write("<br><br>\n");
          f.write("<ul>\n");

          for messageId in Messages:
               if Messages[messageId].messageYear != year:
                    continue
               messageSubject = Messages[messageId].messageSubject
               messageSender = Messages[messageId].messageSender
               messageSenderName = Messages[messageId].messageSenderName
               messageTimeStamp = Messages[messageId].messageTimeStamp
               messageDateTime = datetime.fromtimestamp(float(messageTimeStamp)).strftime("%b %-d, %Y")
               f.write(" <li><a name='{}'></a><a href='{}.html'>{}</a>, <em>{}</em> ({})\n".format(messageId, messageId, cgi.escape(messageSubject), cgi.escape(messageSenderName), messageDateTime));
                    
          f.write("</ul>\n");

          f.write("<br><br>Generated by the <a href='https://github.com/andrei-radulescu-banu/YahooGroups-Archiver'>YahooGroups-Archiver</a> on {}.".format(datetime.now().strftime("%b %-d, %Y")))

          f.close()

          print("Yahoo Thread archived to date-index-{}.html".format(year))

     except Exception as e:
          print("Yahoo Message: {} had an error: {}".format(archiveDateFile, e))

def archiveYahooBySender(messageSenderIndexFileName, archiveDir, format):
     try:
          archiveSenderFile = "{}/sender-index-{}.html".format(archiveDir, messageSenderIndexFileName)

          messageId = Senders[messageSenderIndexFileName].headMessageId
          messageSenderName = Messages[messageId].messageSenderName

          # Update the archive file
          f = open(archiveSenderFile, "w")
          f.write("<style>pre {white-space: pre-wrap;}</style>\n");
          f.write("<h3>{}@yahoogroups.com messages from <em>{}</em></h3>\n".format(groupName, messageSenderName));

          f.write("<ul>\n");

          while messageId:
               messageSubject = Messages[messageId].messageSubject
               messageSender = Messages[messageId].messageSender
               messageSenderName = Messages[messageId].messageSenderName
               messageTimeStamp = Messages[messageId].messageTimeStamp
               messageDateTime = datetime.fromtimestamp(float(messageTimeStamp)).strftime("%b %-d, %Y")
               f.write(" <li><a name='{}'></a><a href='{}.html'>{}</a>, <em>{}</em> ({})\n".format(messageId, messageId, cgi.escape(messageSubject), cgi.escape(messageSenderName), messageDateTime));

               messageId = Messages[messageId].messageSenderNext
          f.write("</ul>\n");

          f.write("<br><br>Generated by the <a href='https://github.com/andrei-radulescu-banu/YahooGroups-Archiver'>YahooGroups-Archiver</a> on {}.".format(datetime.now().strftime("%b %-d, %Y")))

          f.close()

          print("Yahoo Thread archived to sender-index-{}.html".format(messageSenderIndexFileName))

     except Exception as e:
          print("Yahoo Message: {} had an error: {}".format(archiveSenderFile, e))
          
def archiveYahooIndex(archiveDir, format):
     try:
          archiveIndexFile = "{}/index.html".format(archiveDir)

          # Update the archive file
          f = open(archiveIndexFile, "w")
          f.write("<style>pre {white-space: pre-wrap;}</style>\n");
          f.write("<h3>The {}@yahoogroups.com Archive</h3>\n".format(groupName));
          f.write("<ul>\n");

          for year in Threads:
               f.write(" <li><a href='thread-index-{}.html'>Threads in {}</a>\n".format(year, year))

          f.write("</ul>\n");
          f.write("<ul>\n");

          for year in Threads:
               f.write(" <li><a href='date-index-{}.html'>By date in {}</a>\n".format(year, year))
               
          f.write("</ul>\n");

          f.write("<br>Generated by the <a href='https://github.com/andrei-radulescu-banu/YahooGroups-Archiver'>YahooGroups-Archiver</a> on {}".format(datetime.now().strftime("%b %-d, %Y")))
          
          f.close()
          
          print("Yahoo Index archived to index.html")

     except Exception as e:
          print("Archive index had an error: {}".format(archiveIndexFile, e))
    
def loadYahooMessage(fileName, format):
    f1 = open(fileName,"r")
    fileContents=f1.read()
    f1.close()
    jsonDoc = json.loads(fileContents)
    
    if "ygData" not in jsonDoc:
         return None
    
    messageId = jsonDoc["ygData"]["msgId"]
    messageSender = HTMLParser.HTMLParser().unescape(jsonDoc["ygData"]["from"]).decode(format).encode("utf-8")
    messageTimeStamp = jsonDoc["ygData"]["postDate"]
    messageDateTime = datetime.fromtimestamp(float(messageTimeStamp)).strftime("%b %-d, %Y %-I:%M %p")
    messageSubject = HTMLParser.HTMLParser().unescape(jsonDoc["ygData"]["subject"]).decode(format).encode("utf-8")
    messageString = HTMLParser.HTMLParser().unescape(jsonDoc["ygData"]["rawEmail"]).decode(format).encode("utf-8")
    message = email.message_from_string(messageString)
    messageBody = getEmailBody(message)

    messageYear = Messages[messageId].messageYear
    messageDatePrev = Messages[messageId].messageDatePrev
    messageDateNext = Messages[messageId].messageDateNext
    messageThread = Messages[messageId].messageThread
    messageThreadPrev = Messages[messageId].messageThreadPrev
    messageThreadNext = Messages[messageId].messageThreadNext
    messageSenderIndexFileName = Messages[messageId].messageSenderIndexFileName
    
    messageText =  ""
    if messageDatePrev:
         messageText += "[<a href='{}.html'>Date prev</a>]".format(messageDatePrev)
    else:
         messageText += "[Date prev]"
    if messageDateNext:
         messageText += "[<a href='{}.html'>Date next</a>]".format(messageDateNext)
    else:
         messageText += "[Date next]"
    if messageThreadPrev:
         messageText += "[<a href='{}.html'>Thread prev</a>]".format(messageThreadPrev)
    else:
         messageText += "[Thread prev]"
    if messageThreadNext:
         messageText += "[<a href='{}.html'>Thread next</a>]".format(messageThreadNext)
    else:
         messageText += "[Thread next]"
    messageText += "[<a href='date-index-{}.html#{}'>Date index</a>]".format(messageYear, messageId)
    messageText += "[<a href='thread-index-{}.html#{}'>Thread index</a>]".format(messageYear, messageThread)
    messageText += "<br><br>\n"
    messageText += "<font color='#0033cc'>\n"
    messageText += "Sender: <a href='sender-index-{}.html'>{}</a> ({}) <br>\n".format(messageSenderIndexFileName, cgi.escape(messageSender), cgi.escape(messageDateTime))
    messageText += "Subject: {} <br>\n".format(cgi.escape(messageSubject))
    messageText += "<br>\n"
    messageText += "</font>\n"
    messageText += messageBody
    return messageText
    
def getYahooMessages(fileName, format):
    f1 = open(fileName,"r")
    fileContents=f1.read()
    f1.close()

    try:
         jsonDoc = json.loads(fileContents)
         if "ygData" not in jsonDoc or "postDate" not in jsonDoc["ygData"]:
              return None, None, None, None

         messageId = jsonDoc["ygData"]["msgId"]
         messageSender = HTMLParser.HTMLParser().unescape(jsonDoc["ygData"]["from"]).decode(format).encode("utf-8")
         messageTimeStamp = jsonDoc["ygData"]["postDate"]
         messageSubject = HTMLParser.HTMLParser().unescape(jsonDoc["ygData"]["subject"]).decode(format).encode("utf-8")
         
         return messageId, messageSender, messageTimeStamp, messageSubject
    except Exception as e:
         print("Yahoo Message: {} had an error: {}".format(fileName, e))

    return None, None, None, None

# Thank you to the help in this forum for the bulk of this function
# https://stackoverflow.com/questions/17874360/python-how-to-parse-the-body-from-a-raw-email-given-that-raw-email-does-not
def getEmailBody(message):
    body = ""
    if message.is_multipart():
        for part in message.walk():
            ctype = part.get_content_type()
            cdispo = str(part.get("Content-Disposition"))

            # skip any text/plain (txt) attachments
            if ctype == "text/plain" and "attachment" not in cdispo:
                body += "<pre>"
                body += cgi.escape(part.get_payload(decode=True))  # decode
                body += "</pre>"
                break
    # not multipart - i.e. plain text, no attachments, keeping fingers crossed
    else:
        ctype = message.get_content_type()
        if ctype != "text/html":
             body += "<pre>"
             body += cgi.escape(message.get_payload(decode=True))
             body += "</pre>"
        else:
             body += message.get_payload(decode=True)
    return body

## This is where the script starts

parser = argparse.ArgumentParser()

parser.add_argument("-g", "--group", help="Name of Yahoo group", required=True)
parser.add_argument("--skip-year", help="Skip year", action='append', nargs='+')

args = parser.parse_args()

groupName = args.group
skipYears = args.skip_year

oldDir = os.getcwd()
if os.path.exists(groupName):
    archiveDir = os.path.abspath(groupName + "-archive")
    if not os.path.exists(archiveDir):
         os.makedirs(archiveDir)
    os.chdir(groupName)

    # Track the previous message id
    prevMessageId = None

    for fileName in natsorted(os.listdir(os.getcwd())):
         messageId, messageSender, messageTimeStamp, messageSubject = getYahooMessages(fileName, "utf-8")
         if not messageId or not messageSender or not messageTimeStamp or not messageSubject:
              print("Yahoo Message: {} had an error (messageId={}, messageSender={}, messageTimeStamp={}, messageSubject={})".format(fileName, messageId, messageSender, messageTimeStamp, messageSubject))
              continue
         
         messageYear = datetime.fromtimestamp(float(messageTimeStamp)).year
         skip = False
         for skipYearsGroup in skipYears:
              if str(messageYear) in skipYearsGroup:
                   print("Yahoo Message: {} skipped, message year {}".format(fileName, messageYear))
                   skip = True
                   break

         if skip:
              continue
         
         # Save the message metadata
         Messages[messageId] = MessageData()
         Messages[messageId].messageSender = messageSender
         Messages[messageId].messageSenderName = senderName(messageSender)
         Messages[messageId].messageSenderNext = None
         Messages[messageId].messageSenderPrev = None
         Messages[messageId].messageSenderIndexFileName = senderIndexFileName(Messages[messageId].messageSenderName)
         Messages[messageId].messageTimeStamp = messageTimeStamp
         Messages[messageId].messageYear = messageYear
         Messages[messageId].messageSubject = messageSubject
         Messages[messageId].messageThread = messageId
         Messages[messageId].messageThreadNext = None
         Messages[messageId].messageThreadPrev = None
         Messages[messageId].messageDateNext = None
         Messages[messageId].messageDatePrev = prevMessageId
         
         #
         # Thread chaining
         #
         if messageYear not in Threads:
              Threads[messageYear] = OrderedDict()

         for threadId in Threads[messageYear]:
              if Threads[messageYear][threadId].messageSubject in messageSubject:
                   # Chain this message at the end of the thread
                   Messages[messageId].messageThread = threadId
                   tailId = Threads[messageYear][threadId].tailId
                   Messages[messageId].messageThreadPrev = tailId
                   Messages[tailId].messageThreadNext = messageId
                   Threads[messageYear][threadId].tailId = messageId
                   break

         if Messages[messageId].messageThread == messageId:
              # Create new thread for this message
              Threads[messageYear][messageId] = ThreadData()
              Threads[messageYear][messageId].messageSubject = messageSubject
              Threads[messageYear][messageId].tailId = messageId
         
         #
         # Date chaining
         #
         if prevMessageId:
              Messages[prevMessageId].messageDateNext = messageId

         # Update the prev message id
         prevMessageId = messageId        

         #
         # Author chaining
         #
         sn = Messages[messageId].messageSenderIndexFileName

         if sn not in Senders:
              Senders[sn] = SenderData()
              Senders[sn].headMessageId = messageId
         else:
              Messages[Senders[sn].tailMessageId].messageSenderNext = messageId
              Messages[messageId].messageSenderPrev = Senders[sn].tailMessageId
              
         Senders[sn].tailMessageId = messageId
              
              
    for messageId in Messages:
         archiveYahooMessage(messageId, archiveDir, "utf-8")

    for year in Threads:
         archiveYahooByThread(year, archiveDir, "utf8")
         archiveYahooByDate(year, archiveDir, "utf8")

    for sn in Senders:
         archiveYahooBySender(sn, archiveDir, "utf8")

    archiveYahooIndex(archiveDir, "utf8")
         
else:
    sys.exit("Please run archive-group.py first")

os.chdir(oldDir)
print("Complete")


