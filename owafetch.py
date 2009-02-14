#!/usr/bin/env python
# ----------------------------------------------------------------------------
# Copyright (c) 2009 Pieter Kitslaar
# All rights reserved.
# 
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions 
# are met:
#
#  * Redistributions of source code must retain the above copyright
#    notice, this list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright 
#    notice, this list of conditions and the following disclaimer in
#    the documentation and/or other materials provided with the
#    distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
# "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
# LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
# FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
# COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
# INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
# BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT
# LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
# ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.
# ----------------------------------------------------------------------------

#! /usr/bin/env python
"""
Command line utility to fetch mail from a MS Exchange server using WebDAV with Outlook Web Access (OWA) enabled.
Based on the java fetchExc utility (http://www.saunalahti.fi/juhrauti/index.html).
Requires a fetchExc style properties file as input.

Uses the OWAConnection class in the owalib.py file.

Author: Pieter Kitslaar (c) 2009
"""

import sys
import logging
import owalib
import smtplib

def sendMail(_server, _dstAddress, _fromAddress, _message, _user = None, _password = None):
    """
    Sends an email using SMTP.
    """
    s = smtplib.SMTP(_server)
    s.starttls()
    if _user and _password:
        s.login(_user, _password)
    s.sendmail(_fromAddress, _dstAddress, _message)
    s.quit()

# List of required properties 
REQUIRED_PROPERTIES = [
    'ExchangeServer', 
    'Username',
    'Password',
    'Domain',
    'ExchangePath'
    ]

def ParseProperties(_file):
    """
    Parse a fetchExc style properties file.
    """
    # open the file
    f = open(_file)
    
    # define the return dict
    prop_dict = {}

    # loop over the file
    for line in f:
        # split at the '=' character
        key_value = line.split('=')

        # skip if line is not well defined
        if len(key_value) < 2:
            continue
        
        key = key_value[0].strip()
        value = key_value[1].strip()

        # add the key value pair 
        prop_dict[key] = value

    missing_props = [p for p in REQUIRED_PROPERTIES if p not in prop_dict.keys()]
    if len(missing_props) > 0:
        print "Missing following required properties: "
        for p in missing_props:
            print p
        sys.exit(1)

    return prop_dict
        
class FetchExcLoggingFilter(logging.Filter):
    """
    Logging filter class for nice output formatting.
    """
    prefixes = {
                    logging.WARNING: '(W)',
                    logging.DEBUG: '(D)',
                    logging.ERROR: '(E)',
                    logging.INFO: '--',
               }

    def filter(self, record):
        # define the prefix to add to each line
        prefix = self.prefixes.get(record.levelno, "")
        # split the string at line breaks and add prefix to each new line
        record.msg = "\n".join(["%s %s" % (prefix, line) for line in record.msg.split('\n')])
        return True


if __name__ == "__main__":
    # import the command line option parser
    from optparse import OptionParser

    # define the options parser options
    usage = "usage: %prog [options] <properties-file>\n"
    parser = OptionParser(usage)
    parser.add_option("-s", "--silent", action="store_true", dest="Silent", help="Silent outpu. Only outputs error messaghes.")
    parser.add_option("-l", "--list", action="store_true", dest="ListOnly", help="Only list the messages.")
    parser.add_option("-a", "--all", action="store_true", dest="AllMessages", help="Fetch all messages (default: only unread messages)")
    parser.add_option("-v", "--verbose", action="store_true", dest="Verbose", help="Produce verbose output")
    
    # parse the options
    (options, args) = parser.parse_args()

    # make sure a properties-files is given
    if len(args) < 1:
        parser.print_help()
        sys.exit(1)
    
    # get the properties file and parse it
    properties_file = args[0]
    prop_dict = ParseProperties(properties_file)

    # Setup console output for the logging module
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    if options.Silent:
        console.setLevel(logging.ERROR)
    console.addFilter(FetchExcLoggingFilter())
    if options.Verbose:
        console.setLevel(logging.DEBUG)
    console.setFormatter(logging.Formatter('%(message)s'))

    log = logging.getLogger('FetchExc')
    log.setLevel(logging.DEBUG)
    log.addHandler(console)

    # get the full username with the domain part
    fullUserName = "%s\\%s" % (prop_dict["Domain"], prop_dict["Username"])

    # Create the OWA session
    owa  = owalib.OWAConnection(prop_dict["ExchangeServer"])

    # Use Form Based Authentication (FBA)
    owa.doFBA( fullUserName, prop_dict["Password"], prop_dict["ExchangePath"], prop_dict["FBApath"] )

    # get the users root path on the server
    rootpath = owa.getRootPath(prop_dict["ExchangePath"])
    log.info('Found user root path: %s' % rootpath)

    # get the users inbox path
    inboxPath = owa.getInboxPath(rootpath)
    log.info('Found inbox path: %s' % inboxPath)


    # see which messages to list (default only unread)
    if options.AllMessages or prop_dict.get("All",'false').lower() == 'true':
        message_type = ""
        fetchAll = True
    else:
        message_type = " unread"
        fetchAll = False

    # get a list of the messages in the inbox
    inboxMessages =  owa.getListMessages(inboxPath, fetchAll)
    log.info('Found %i%s message(s).' % (len(inboxMessages), message_type) )
    
    # loop over the messages
    for i, m in enumerate(inboxMessages):
        # only print message info 
        if options.ListOnly:
            log.info('(%i) (%s)%s' % (i, m["fromemail"], m["subject"]))
        else:
            log.info('Sending mail %i: (%s)%s' % (i, m["fromemail"], m["subject"]))
            message = owa.getMessage(m["href"])

            sendMail(prop_dict["MailServer"], prop_dict["DestinationAddress"], m["fromemail"],message)
            
            if prop_dict.get('Delete','false').lower() == 'true':
                if(owa.deleteMessage(inboxPath, m["href"])):
                    log.info('Deleted message %i.' % i)
            else:
                if(owa.markAsRead(inboxPath, m["href"])):
                    log.info('Marked message %i read.' % i)


    owa.close()
