#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo, Locale
from com.sun.star.task import XJobExecutor, XInteractionHandler
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


class PyOAuth2Service(unohelper.Base, XServiceInfo, XJobExecutor, XDialogEventHandler, XStringEscape, XInteractionHandler):
    def __init__(self, ctx):
        self.ctx = ctx
        self.dialog = None
        self.configuration = "com.gmail.prrvchr.extensions.gMailOOo/MailMergeWizard"
        self.redirect = "urn:ietf:wg:oauth:2.0:oob"
        self.secret = str.encode(str(uuid.uuid4().hex + uuid.uuid4().hex))
        self._loadAvencedConfig()
        resource = "com.sun.star.resource.StringResourceWithLocation"
        arguments = (self._getResourceLocation(), True, self._getCurrentLocale(), "DialogStrings", "", self)
        self.resource = self.ctx.ServiceManager.createInstanceWithArgumentsAndContext(resource, arguments, self.ctx)

    # XJobExecutor
    def trigger(self, step=3):
        try:
            self._openDialog(int(step))
        except Exception as e:
            print(e)

    # XInteractionHandler
    def handle(self, requester):
        pass

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        if self.dialog == dialog:
            step = dialog.Model.Step
            if method == "DialogOk":
                if step == 1:
                    self._saveAvencedConfig()
                    self._loadAvencedConfig()
                    self._setDialogStep(2)
                elif step == 2:
                    self._executeShell(self._getAuthorizationUrl())
                    self._setDialogStep(3)
                elif step == 3:
                    code = dialog.getControl("TextField6").getText()
                    if code:
                        self._getTokens(code)
                    dialog.endExecute()
                return True
            elif method == "DialogCancel":
                dialog.endExecute()
                return True
            elif method == "Settings":
                self._setDialogStep(1)
                return True
        return False
    def getSupportedMethodNames(self):
        return ("DialogOk", "DialogCancel", "Settings")

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

    def _openDialog(self, step=3):
        provider = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.DialogProvider", self.ctx)
        self.dialog = provider.createDialogWithHandler("vnd.sun.star.script:gMailOOo.Dialog?location=application", self)
        self._setDialogStep(step)
        self.dialog.execute()
        self.dialog.dispose()
        self.dialog = None

    def _setDialogStep(self, step):
        if step == 1:
            self.dialog.Title = self.resource.resolveString("1.Dialog.Title")
            configuration = self._getConfiguration(self.configuration)
            self.dialog.getControl("TextField1").setText(configuration.getByName("ClientId"))
            self.dialog.getControl("TextField2").setText(configuration.getByName("AuthorizationUrl"))
            self.dialog.getControl("TextField3").setText(configuration.getByName("TokenUrl"))
            self.dialog.getControl("TextField4").setText(configuration.getByName("Scope"))
        elif step == 2:
            self.dialog.Title = self.resource.resolveString("2.Dialog.Title")
            self.dialog.getControl("TextField5").setText(self._getAuthorizationUrl())
        elif step == 3:
            self.dialog.Title = self.resource.resolveString("3.Dialog.Title")
        self.dialog.Model.Step = step

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

    def _getAuthorizationUrl(self):
        params = {}
        params["client_id"] = self.clientId
        params["redirect_uri"] = self.redirect
        params["response_type"] = "code"
        params["scope"] = self.scope
        params["access_type"] = "offline"
        params["code_challenge_method"] = "S256"
        params["code_challenge"] = self._getChallengeCode()
        params["login_hint"] = self._getUserName()
        return requests.Request("GET", self.authorizationUrl, params=params).prepare().url

    def _getChallengeCode(self):
        code = hashlib.sha256(self.secret).digest()
        padding = {0:0, 1:2, 2:1}[len(code) % 3]
        challenge = base64.urlsafe_b64encode(code)
        return challenge[:len(challenge)-padding]

    def _getTokens(self, code):
        timestamp = int(time.time())
        data = {}
        data["client_id"] = self.clientId
        data["redirect_uri"] = self.redirect
        data["grant_type"] = "authorization_code"
        data["code"] = code
        data["code_verifier"] = self.secret
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        response = requests.post(self.tokenUrl, headers=headers, data=data, timeout=5)
        self._saveResponse(response.json(), timestamp)

    def _isAccessTokenExpired(self, timestamp):
        return self._getConfiguration(self.configuration).getByName("ExpiresTimeStamp") <= timestamp

    def _saveAvencedConfig(self):
        configuration = self._getConfiguration(self.configuration, True)
        configuration.replaceByName("ClientId", self.dialog.getControl("TextField1").getText())
        configuration.replaceByName("AuthorizationUrl", self.dialog.getControl("TextField2").getText())
        configuration.replaceByName("TokenUrl", self.dialog.getControl("TextField3").getText())
        configuration.replaceByName("Scope", self.dialog.getControl("TextField4").getText())
        configuration.commitChanges()

    def _loadAvencedConfig(self):
        self.clientId = self._getConfiguration(self.configuration).getByName("ClientId")
        self.authorizationUrl = self._getConfiguration(self.configuration).getByName("AuthorizationUrl")
        self.tokenUrl =  self._getConfiguration(self.configuration).getByName("TokenUrl")
        self.scope = self._getConfiguration(self.configuration).getByName("Scope")

    def _getAuthorizationCode(self):
        self._executeShell(self._getAuthorizationUrl())
        self._openDialog(2)
        return self._getConfiguration(self.configuration).getByName("RefreshToken")

    def _refreshAccessToken(self, refreshtoken, timestamp):
        data = {}
        data["client_id"] = self.clientId
        data["refresh_token"] = refreshtoken
        data["grant_type"] = "refresh_token"
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        response = requests.post(self.tokenUrl, headers=headers, data=data, timeout=5)
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
                                         "com.sun.star.task.XJobExecutor",
                                         "com.sun.star.awt.XDialogEventHandler",
                                         "com.sun.star.util.XStringEscape",
                                         "com.sun.star.task.XInteractionHandler"), )                 # List of implemented services
