## 1. Create a properties file ##
The owafetch utility requires a properties file as input. The syntax and variables in this properties file are identical to the [fetchExc](http://www.saunalahti.fi/juhrauti/index.html) properties file (for easy migration from fetchExc to owafetch). See comments in the [example.propeties](http://owalib.googlecode.com/svn/trunk/example.properties) file for more information about the available properties.
Below is an example of a properties file.

```
ExchangeServer = xxx.dddddd.com
ExchangePath = exchange
Username = domainuser 
Password = domainpassword
Domain = DOMAIN
Secure = True
DestinationAddress = user@yyy.dddddd.com
```

**N.B. Don't use the 'Delete = True' options until you are sure that all your other settings are correct.**


## 2. Test the Outlook Web Access connection ##
Check the settings for accessing your MS Exchange Server through Outlook Web Access by listing (**-l** or **--list**) the (unread) messages on your Inbox.
```
> owafetch.py example.prefs -l
-- Found user root path: https://xxx.dddddd.com/exchange/domainuser/
-- Found inbox path: https://xxx.dddddd.com/exchange/domainuser/Inbox
-- Found 3 unread message(s).
-- (0) (sales@s***.com)Buy stuff
-- (1) (support@b**.com)RE: your question
-- (2) (john.doe@h**mail.com) Drinks?
```

If you have no unread messages in your inbox you can use the **-a** (or **--all**) option to list all messages. This could take a couple of seconds if you have a lot of messages in your Inbox.

```
> owafetch.py example.prefs -l -a
-- Found user root path: https://xxx.dddddd.com/exchange/domainuser/
-- Found inbox path: https://xxx.dddddd.com/exchange/domainuser/Inbox
-- Found 1438 message(s).
-- (0) (foo@bar**.com)Meeting at 9pm
etc..
```

## 3. Test the mail forwarding ##
Test the forwarding of the fetched messages by calling owafetch without the -l option.
The typical result will look like this.
```
> owafetch.py example.prefs
-- Found user root path: https://xxx.dddddd.com/exchange/domainuser/
-- Found inbox path: https://xxx.dddddd.com/exchange/domainuser/Inbox
-- Found 1 unread message(s).
-- Sending mail 0: (john@server**.net)New features
-- Marked message 0 read.
```

After the message has been succesfully send to the destination address it will be marked as 'read' on the MS Exchange server.

## 4. Automate e-mail fetching/forwarding ##
If you have confirmed that your settings are correct you can automate the calling of owafetch (for example using `cron`). You can use the **-s** (or **--silent**) option to silence the output of the program (except for error messages).

Below is an example setting for `cron` which fetches mail from the server every 5 minutes on working days (Monday till Friday) and once every hour during Saturday and Sunday.

```
*/5 * * * 1-5 /home/foo/owalib/owafetch.py -s ~/.owafetch/MyCompany.properties
0 */1 * * 0,6 /home/foo/owalib/owafetch.py -s ~/.owafetch/MyCompany.properties
```