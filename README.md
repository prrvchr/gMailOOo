# gMailOOo



Google Mail OAuth2 implementation for LibreOffice.



![alt text](gMailOOo.png "gMailOOo screenshot")



Has been tested with:
	
LibreOffice 5.3.3.2 - Lubuntu 16.10 -  LxQt 0.11.0.3

LibreOffice 5.3.1.2 x86 - Windows 7 SP1


Gmail account setting: 

Smtp Server: smtp.gmail.com

User: your gMail email address (mandatory)

Password: your gMail password (needed for user/password authentication)


Type of connection tested:

SSL on port 465

TLS on port 587 (recommanded connection type)

OAuth2 with SSL or TLS


Setting requirement for user/password authentication with SSL or TLS:

Load: https://myaccount.google.com/security?utm_source=OGB#connectedapps

You must enable less secured application.
