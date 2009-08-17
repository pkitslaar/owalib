#! /usr/bin/env python
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

"""
Classes and methods to access a MS Exchange server through WebDAV with Outlook Web Access (OWA)

Inspired by and based on:
 - fetchExc (java): http://www.saunalahti.fi/juhrauti/index.html
 - davlib.py (python): http://www.webdav.org/mod_dav/davlib.py
 - http://pxa-be.blogspot.com/2008/07/exchange-form-based-authentication-and.html

Author: Pieter Kitslaar (c) 2009
"""

import httplib
import urllib2
import cookielib
import sys
import re

XML_CONTENT_TYPE = 'text/xml; charset="utf-8"'

def getDeleteMsg(_filename): 
    """
    XML code for a DELETE request to delete a message (no copy to "Trash", a real delete).
    Should be send to the "inbox", but _filename is the full href of the message.
    """

    strBuf = """
    <?xml version="1.0" encoding="utf-8" ?>
    <D:delete xmlns:D="DAV:" xmlns:a="urn:schemas:httpmail:">
        <D:target><D:href>%s</D:href></D:target>
    </D:delete>
    """
    return strBuf % _filename.split('/')[-1]


def getMarkAsReadMsg(_filename):
    """
    XML code for a BPROPPATCH request to mark a message as read.
    Should be send to the "inbox", but _filename is the full href of the message.
    """ 

    strBuf = """ 
    <?xml version="1.0" encoding="utf-8" ?>
    <D:propertyupdate xmlns:D="DAV:" xmlns:a="urn:schemas:httpmail:">
      <D:target><D:href>%s</D:href></D:target>
      <D:set><D:prop><a:read>1</a:read></D:prop></D:set>
    </D:propertyupdate>
    """
    return strBuf % _filename.split('/')[-1]

def getInboxMsg(): 
    """
    XML code for  a PROPFIND request to obtain the "inbox" property.
    """ 

    strBuf = """
    <?xml version="1.0" encoding="utf-8" ?>
    <D:propfind xmlns:D="DAV:" xmlns:a="urn:schemas:httpmail:">
      <D:prop>
        <a:inbox/>
      </D:prop>
    </D:propfind>"""
    return strBuf

def getListMailMsg(_allMsgs = False): 
    """
    XML code for a SEARCH request to get all the messages in a folder.
    If _allMsgs == True it will return all messages. Else only the unread messages.
    """

    strBuf = """
    <?xml version="1.0" encoding="utf-8" ?>
    <searchrequest xmlns="DAV:">
      <sql>
        SELECT 
            "urn:schemas:httpmail:fromemail", 
            "urn:schemas:httpmail:subject", 
            "urn:schemas:httpmail:read" 
        FROM ""
        WHERE &quot;DAV:iscollection&quot; = False AND &quot;DAV:ishidden&quot; = False
        %s
        ORDER BY "DAV:creationdate"
      </sql>
    </searchrequest>"""

    and_statement = ""
    if not _allMsgs:
        and_statement = """ AND "urn:schemas:httpmail:read"= False"""
    return strBuf % and_statement


# list of characters with their substitute HEX code for nicer href strings
FIX_REPLACE = []
FIX_REPLACE.append( ("[", "%5B") )
FIX_REPLACE.append( ("]", "%5D") )
FIX_REPLACE.append( ("|", "%7C") )
FIX_REPLACE.append( ("^", "%5E") )
FIX_REPLACE.append( ("`", "%60") )
FIX_REPLACE.append( ("{", "%7B") )
FIX_REPLACE.append( ("}", "%7D") )


def fixFileName(_fileName):
    """
    Replaces unwanted characters in returned href strings from the Echange server with there hex equivalents.
    """

    result = _fileName
    for repl in FIX_REPLACE:
        result = result.replace(repl[0], repl[1])
    return result 


class OWAConnectionPlugin(object):
    """
    Plugin class to connect to a Outlook Web Access page using WebDAV.
    Don't use this class directly, but use the specialized classes:
        - PlainOWAConnection
        - SecureOWAConnection
    Or use the OWAConnectionClass factory method.
    """
    def __init__(self, *args, **kw):
        # using super for mixin multiple inheritance
        # as discussed at: http://stackoverflow.com/questions/1282368/
        s  = super(OWAConnectionPlugin, self) 
        s.__init__(*args, **kw)
        
        self.authentication_header = None

    def doFBA(self, _username, _password, _exchangePath, _fbaPath):
        """
        Try to authenticate with form based authentication (FBA).
        Stores the authentication cookie in the authentication_header.

        This is based on (copied from) the example found in Phil Andrews blog:
        http://pxa-be.blogspot.com/2008/07/exchange-form-based-authentication-and.html
        """
        
        # init the CookieJar and register it with the url openner
        cj = cookielib.CookieJar()
        opener = urllib2.build_opener(urllib2.HTTPSHandler(),urllib2.HTTPCookieProcessor(cj))
        urllib2.install_opener(opener)

        # define the body of the request setting the username and password
        owabody = ''
        owabody += 'destination=https://' + self.host + '/' + _exchangePath + '/'
        owabody += '&username=' + _username
        owabody += '&password=' + _password

        # the headers
        # TODO: Check if the User-Agent entry is needed
        owaheaders = {'Content-Type': 'application/x-www-form-urlencoded',
        'Connection': 'Keep-Alive',
        'User-Agent': 'Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.1.4322)',
        'Host': self.host}

        # define the url and make the request
        owaurl = 'https://'+self.host+_fbaPath
        owareq = urllib2.Request(owaurl,owabody,owaheaders)
        owa = urllib2.urlopen(owareq)

        # get the cookie from the jar
        cookie_string = ''
        for i in cj:
            cookie_string += i.name+'='+i.value+';'

        # add the cookie to the authentication_header
        # this will be used in subsequent calls to the server
        self.authentication_header  = {'Cookie': cookie_string}


    def do_request(self, method, url, body=None, extra_hdrs={}):
        """
        Make a request to the server.
        """
        request_header = extra_hdrs.copy()
        if self.authentication_header != None:
            request_header.update(self.authentication_header)
        self.request(method, url, body, request_header)
        return self.getresponse()
    
    def getRootPath(self, _pRootPath):
        """
        Return the user's root path on the server based on the given exchange path.
        """
        resp = self.do_request('GET', "/" + _pRootPath + "/")
        status = resp.status
        if status == 401:
            #self.log.error("Authentication failed. Exiting!")
            sys.exit(1)
        elif status == 404:
            #self.log.error("Exchange user path doesn't exist. Exiting!")
            sys.exit(1)
        elif status == 200:
            pass
        else:
            #self.log.error("Status = %s. Exiting!" % status)
            sys.exit(1)

        # get the respons text
        text = resp.read()

        # find the content of the <BASE href="..."> element
        #                                     ^^^ 
        a = re.compile('<BASE href="([^"]+)">')
        m = a.search(text)

        if not m:
            self.log.error("Could not find <BASE href=\"..\"> tag. Exiting!")
            sys.exit(1)
        
        return m.group(1)

    def getInboxPath(self, _rootPath):
        """
        Get the inbox path based on the given root path.
        """
        # create additional header dictionary
        hrd = {}
        hrd['Depth'] = '1'
        hrd['Content-Type'] = XML_CONTENT_TYPE
        resp = self.do_request('PROPFIND', _rootPath, getInboxMsg(), hrd)

        if resp.status != 207:
            #self.log.error("Wrong stats (%s). Exiting!" % resp.status)
            sys.exit(1)

        xml  = resp.read()

        a = re.compile(r'<d:inbox>([^<]+)</d:inbox>')
        m = a.search(xml)

        if not m:
            #self.log.error("Could not find inbox path. Exiting!")
            sys.exit(1)
        
        return m.groups()[0]

    def getListMessages(self, _inboxPath, _all = False):
        """
        Returns the url, subject and fromemail of the unread (or all) messages in the _inboxPath.
        Returns a list of dictionaries with keys: 'href', 'fromemail' and 'subject'
        """
        hrd = {'Depth': '1', 'Content-Type': XML_CONTENT_TYPE }
        resp = self.do_request('SEARCH', _inboxPath, getListMailMsg(_all), hrd)
        xml =  resp.read()

        # regexp to find each response body
        a = re.compile(r'<a:response>(.+?)</a:response>')
        # regeexp to find the "href", "fromemail" and "subject" contents
        b = re.compile(r'<a:href>(?P<href>.+?)</a:href>.*?<d:fromemail>(?P<fromemail>.+?)</d:fromemail>.*?<d:subject>(?P<subject>.*?)</d:subject>')

        raw_list =  a.findall(xml)

        messages = []
        for item in raw_list:
            m = b.search(item)
            if not m:
                print "Could not match regexp in: %s" % item
                continue
            mail_dict = {}
            mail_dict.update(m.groupdict())
            messages.append( mail_dict )

        return messages

    def getMessage(self, _messagePath):
        """
        Return the raw text of the message.
        """
        hdr = {'Translate': 'F'}
        resp = self.do_request('GET', _messagePath, "", hdr)
        return resp.read()

    def markAsRead(self, _inboxPath, _messagePath):
        """
        Mark a message as read. Requires the path to the "inbox" and the full path to the message.
        """
        hdr = {'Content-type': XML_CONTENT_TYPE}
        resp = self.do_request('BPROPPATCH', _inboxPath + '/', getMarkAsReadMsg(_messagePath), hdr)
        resp.read() # read to make sure we can do another request
        return resp.status == 207

    def deleteMessage(self, _inboxPath, _messagePath):
        """
        Delete a message on the server. 
        N.B. Does not move the message to "Trash" really deletes it!

        Requires the path to the "inbox" and the full path to the message.
        """
        hdr = {'Content-type': XML_CONTENT_TYPE}
        resp= self.do_request('BDELETE', _inboxPath + '/', getDeleteMsg(_messagePath), hdr)
        resp.read()
        return resp.status == 207


class PlainOWAConnection(OWAConnectionPlugin, httplib.HTTPConnection, object):
    """ OWAConnection class for non-SSL (http://) connections """
    pass
    


class SecureOWAConnection(OWAConnectionPlugin, httplib.HTTPSConnection, object):
    """ OWAConnection class for secure SSL (https://) connections """
    pass

def OWAConnectionClass(_secure):
    """ Factory method to return the correc connection class based on the _secure argument. """
    if _secure:
        return SecureOWAConnection
    else:
        return PlainOWAConnection

    
