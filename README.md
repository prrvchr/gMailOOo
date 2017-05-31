# gMailOOo

## Google Mail OAuth2.0 implementation for LibreOffice.

![gMailOOo screenshot](gMailOOo.png)

## Has been tested with:
	
* LibreOffice 5.3.3.2 - Lubuntu 16.10 -  LxQt 0.11.0.3

* LibreOffice 5.3.1.2 x86 - Windows 7 SP1

## Gmail account setting: 

* Smtp Server: smtp.gmail.com

* User: your gMail email address (mandatory)

* Password: your gMail password (needed for user/password authentication, not used with OAuth2 but do not leave empty)

## Type of connection tested:

* SSL on port 465

* TLS on port 587 (recommanded connection type)

## Type of authentication tested:

* Login/password  with SSL or TLS

* OAuth2 with SSL or TLS (recommanded authentication type)

## Setting requirement for Login/password authentication with SSL or TLS:

Load [Google Account Setting](https://myaccount.google.com/security?utm_source=OGB#connectedapps)

You must enable less secured application.

## Setting requirement for OAuth2 authentication with SSL or TLS:

You must get authorization code from Google:

LibreOffice mailmerge wants to Send email on your behalf for sending email

copy and paste authorization code to LibreOffice Message Box.
