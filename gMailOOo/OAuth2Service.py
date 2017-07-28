#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo, XInitialization
from com.sun.star.task import XInteractionHandler
from com.sun.star.awt import XDialogEventHandler
from com.sun.star.util import XStringEscape
from com.sun.star.beans import PropertyValue

import requests
import time
import uuid
import base64
import hashlib

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = "com.gmail.prrvchr.extensions.gMailOOo.OAuth2Service"


class PyOAuth2Service(unohelper.Base, XServiceInfo, XInitialization, XDialogEventHandler, XStringEscape, XInteractionHandler):
    def __init__(self, ctx):
        self.ctx = ctx
        self.dialog = None
        self.configuration = "com.gmail.prrvchr.extensions.gMailOOo/MailMergeWizard"
        self.client = "85733218570-ijl14qpp5sqq029lg89qvppgm29qhklt.apps.googleusercontent.com"
        self.redirect = "urn:ietf:wg:oauth:2.0:oob"
        self.secret = str.encode(str(uuid.uuid4().hex * 2))
        #resource = "com.sun.star.resource.StringResourceWithLocation"
        #arguments = (self._getResourceLocation(), True, self._getCurrentLocale(), "DialogStrings", "", self)
        #self.resource = self.ctx.ServiceManager.createInstanceWithArgumentsAndContext(resource, arguments, self.ctx)

    # XInitialization
    def initialize(self, *arg):
        url = self._getPermissionUrl()
        self._openDialog(1, url)

    # XInteractionHandler
    def handle(self, requester):
        pass

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        if self.dialog == dialog:
            step = dialog.Model.Step
            if method == "DialogOk":
                if step == 1:
                    url = self._getPermissionUrl()
                    self._executeShell(url)
                    dialog.Model.Step = 2
                elif step == 2:
                    code = dialog.getControl("TextField2").getText()
                    if code:
                        self._getTokens(code)
                    dialog.endExecute()
                return True
            elif method == "DialogCancel":
                dialog.endExecute()
                return True
        return False
    def getSupportedMethodNames(self):
        return ("DialogOk", "DialogCancel")

    # XStringEscape
    def escapeString(self, username):
        return self._getAuthenticationString(username, True)
    def unescapeString(self, username):
        return self._getAuthenticationString(username, False)

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    def _executeShell(self, url, option=""):
        shell = self.ctx.ServiceManager.createInstance("com.sun.star.system.SystemShellExecute")
        shell.execute(url, option, 0)

    def _openDialog(self, step, text=""):
        provider = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.DialogProvider", self.ctx)
        self.dialog = provider.createDialogWithHandler("vnd.sun.star.script:gMailOOo.Dialog?location=application", self)
        self.dialog.Model.Step = step
        if text != "":
            self.dialog.getControl("TextField%s" % (step)).setText(text)
        self.dialog.execute()
        self.dialog.dispose()
        self.dialog = None

    def _getUserName(self):
        path = "/org.openoffice.Office.Writer/MailMergeWizard"
        if self._getMailServiceType() == uno.Enum("com.sun.star.mail.MailServiceType", "SMTP"):
            return self._getConfiguration(path).getByName("MailUserName")
        else:
            return self._getConfiguration(path).getByName("InServerUserName")

    def _getMailServiceType(self):
        path = "/org.openoffice.Office.Writer/MailMergeWizard"
        if self._getConfiguration(path).getByName("IsSMPTAfterPOP"):
            if self._getConfiguration(path).getByName("InServerIsPOP"):
                return uno.Enum("com.sun.star.mail.MailServiceType", "POP3")
            else:
                return uno.Enum("com.sun.star.mail.MailServiceType", "IMAP")
        else:
            return uno.Enum("com.sun.star.mail.MailServiceType", "SMTP")

    def _getPermissionUrl(self, scope="https://www.googleapis.com/auth/gmail.send"):
        parameters = {}
        parameters["client_id"] = self.client
        parameters["redirect_uri"] = self.redirect
        parameters["response_type"] = "code"
        parameters["scope"] = scope
        parameters["access_type"] = "offline"
        parameters["code_challenge_method"] = "S256"
        parameters["code_challenge"] = self._getChallengeCode()
        parameters["login_hint"] = self._getUserName()
        return requests.Request("GET", "https://accounts.google.com/o/oauth2/v2/auth", params=parameters).prepare().url

    def _getChallengeCode(self):
        code = hashlib.sha256(self.secret).digest()
        padding = {0:0, 1:2, 2:1}[len(code) % 3]
        challenge = base64.urlsafe_b64encode(code)
        return challenge[0:len(challenge)-padding]

    def _getAccountsUrl(self, command):
        return "%s/%s" % ("https://www.googleapis.com", command)

    def _getTokens(self, code):
        timestamp = int(time.time())
        parameters = {}
        parameters["client_id"] = self.client
        parameters["redirect_uri"] = self.redirect
        parameters["grant_type"] = "authorization_code"
        parameters["code"] = code
        parameters["code_verifier"] = self.secret
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        response = requests.post(self._getAccountsUrl("oauth2/v4/token"), headers=headers, data=parameters, timeout=5)
        self._saveResponse(response.json(), timestamp)

    def _isAccessTokenExpired(self, timestamp):
        return self._getConfiguration(self.configuration).getByName("ExpiresTimeStamp") <= timestamp

    def _getAuthorizationCode(self):
        url = self._getPermissionUrl()
        self._executeShell(url)
        self._openDialog(2)
        return self._getConfiguration(self.configuration).getByName("RefreshToken")

    def _refreshAccessToken(self, refreshtoken, timestamp):
        parameters = {}
        parameters["client_id"] = self.client
        parameters["refresh_token"] = refreshtoken
        parameters["grant_type"] = "refresh_token"
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        response = requests.post(self._getAccountsUrl("oauth2/v4/token"), headers=headers, data=parameters, timeout=5)
        return self._saveResponse(response.json(), timestamp, "access_token")

    def _getAuthenticationString(self, username, encode=False):
        refreshtoken = self._getConfiguration(self.configuration).getByName("RefreshToken")
        if not refreshtoken:
            refreshtoken = self._getAuthorizationCode()
        timestamp = int(time.time())
        if self._isAccessTokenExpired(timestamp):
            accesstoken = self._refreshAccessToken(refreshtoken, timestamp)
        else:
            accesstoken = self._getConfiguration(self.configuration).getByName("AccessToken")
        authstring = "user=%s\1auth=Bearer %s\1\1" % (username, accesstoken)
        if encode:
          authstring = base64.b64encode(authstring.encode("ascii"))
        return authstring
        
    def _saveResponse(self, response, timestamp, key=None):
        configuration = self._getConfiguration(self.configuration, True)
        if "refresh_token" in response:
            configuration.replaceByName("RefreshToken", response["refresh_token"])
        if "access_token" in response:
            configuration.replaceByName("AccessToken", response["access_token"])
        if "expires_in" in response:
            configuration.replaceByName("ExpiresTimeStamp", timestamp + int(response["expires_in"]))
        configuration.commitChanges()
        return key if key is None else response[key]

    def _getConfiguration(self, nodepath, update=False):
        value = uno.Enum("com.sun.star.beans.PropertyState", "DIRECT_VALUE")
        config = self.ctx.ServiceManager.createInstance("com.sun.star.configuration.ConfigurationProvider")
        service = "com.sun.star.configuration.ConfigurationUpdateAccess" if update else "com.sun.star.configuration.ConfigurationAccess"
        return config.createInstanceWithArguments(service, (PropertyValue("nodepath", -1, nodepath, value),))

    def _getResourceLocation(self):
        identifier = "com.gmail.prrvchr.extensions.gMailOOo"
        provider = self.ctx.getValueByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
        return "%s/gMailOOo" % (provider.getPackageLocation(identifier))

    def _getCurrentLocale(self):
        parts = self._getConfiguration("/org.openoffice.Setup/L10N").getByName("ooLocale").split("-")
        locale = Locale(parts[0], "", "")
        if len(parts) == 2:
            locale.Country = parts[1]
        else:
            service = self.ctx.ServiceManager.createInstance("com.sun.star.i18n.LocaleData")
            locale.Country = service.getLanguageCountryInfo(locale).Country
        return locale


g_ImplementationHelper.addImplementation(PyOAuth2Service,                                            # UNO object class
                                         g_ImplementationName,                                       # Implementation name
                                         ("com.sun.star.lang.XServiceInfo",
                                         "com.sun.star.lang.XInitialization",
                                         "com.sun.star.awt.XDialogEventHandler",
                                         "com.sun.star.util.XStringEscape",
                                         "com.sun.star.task.XInteractionHandler"), )                 # List of implemented services
