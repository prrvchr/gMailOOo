#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo, XInitialization, XLocalizable
from com.sun.star.script.provider import XScript
from com.sun.star.beans import XPropertySet, XPropertyAccess
from com.sun.star.awt import XDialogEventHandler
from com.sun.star.task import XInteractionHandler

import requests
import time
import uuid
import base64
import hashlib

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = "com.gmail.prrvchr.extensions.gMailOOo.OAuth2Service"


class PyOAuth2Service(unohelper.Base, XServiceInfo, XInitialization, XLocalizable, XScript, XPropertySet, XPropertyAccess, XDialogEventHandler, XInteractionHandler):
    def __init__(self, ctx, *namedvalues):
        self.ctx = ctx
        self.dialog = None
        self.dialogurl = "vnd.sun.star.script:gMailOOo.OAuth2Dialog?location=application"
        self.settings = {}
        self.ClientId = ""
        self.AuthorizationUrl = ""
        self.TokenUrl =  ""
        self.Scope = ""
        self.UserName = ""
        self.RefreshToken = ""
        self.AccessToken = ""
        self.TimeStamp = 0
        self.locale = self._getCurrentLocale()
        self.configuration = "com.gmail.prrvchr.extensions.gMailOOo/MailMergeWizard"
        self.redirect = "urn:ietf:wg:oauth:2.0:oob"
        self.secret = str.encode(str(uuid.uuid4().hex + uuid.uuid4().hex))
        self.initialize(namedvalues)

    # XInitialization
    def initialize(self, namedvalues=()):
        self.setPropertyValues(namedvalues)

    # XScript
    def invoke(self, args, outindex, out):
        result = []
        self.setPropertyValues(args)
        if self._getOAuth2Tokens():
            result = self.getPropertyValues() + self.settings.values()
        return tuple(result), None, None

    # XInteractionHandler
    def handle(self, requester):
        pass

    # XLocalizable
    def setLocale(self, locale):
        self.locale = locale
    def getLocale(self):
        return self.locale

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        if self.dialog == dialog:
            if method == "DialogBack":
                self._setDialogStep(1)
                return True
            elif method == "DialogNext":
                if dialog.Model.Step == 1:
                    if self._saveSettings():
                        self._setDialogStep(2)
                else:
                    self._executeShell(dialog.getControl("AuthorizationFullUrl").getText())
                    self._setDialogStep(3)
                return True
        return False
    def getSupportedMethodNames(self):
        return ("DialogBack", "DialogNext")

    # XPropertySet
    def getPropertySetInfo(self):
        return None
    def setPropertyValue(self, name, value):
        if hasattr(self, name):
            setattr(self, name, value)
    def getPropertyValue(self, name):
        if hasattr(self, name):
            return getattr(self, name)
        return None
    def addPropertyChangeListener(self, name, listener):
        pass
    def removePropertyChangeListener(self, name, listener):
        pass
    def addVetoableChangeListener(self, name, listener):
        pass
    def removeVetoableChangeListener(self, name, listener):
        pass

    # XPropertyAccess
    def getPropertyValues(self):
        return (self._getNamedValue("RefreshToken", self.RefreshToken), 
                self._getNamedValue("AccessToken", self.AccessToken),
                self._getNamedValue("TimeStamp", self.TimeStamp))
    def setPropertyValues(self, namedvalues):
        for namedvalue in namedvalues:
            if hasattr(namedvalue, "Name") and hasattr(namedvalue, "Value"):
                self.setPropertyValue(namedvalue.Name, namedvalue.Value)

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    def _getOAuth2Tokens(self):
        timestamp = int(time.time())
        if not self.RefreshToken or not self.AccessToken:
            code = self._getAuthorizationCode()
            if code is None:
                return False
            self._getTokens(code, timestamp)
        elif self.TimeStamp <= timestamp:
            self._refreshAccessToken(timestamp)
        return True

    def _getAuthorizationCode(self):
        code = None
        if self._openDialog():
            code = self.dialog.getControl("AuthorizationCode").getText()
        self.dialog.dispose()
        self.dialog = None
        return code

    def _openDialog(self):
        provider = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.DialogProvider", self.ctx)
        self.dialog = provider.createDialogWithHandler(self.dialogurl, self)
        self.dialog.getControl("ClientId").setText(self.ClientId)
        self.dialog.getControl("AuthorizationUrl").setText(self.AuthorizationUrl)
        self.dialog.getControl("TokenUrl").setText(self.TokenUrl)
        self.dialog.getControl("Scope").setText(self.Scope)
        if self._isInitialized():
            self.dialog.getControl("AuthorizationFullUrl").setText(self._getAuthorizationFullUrl())
            self._setDialogStep(2)
        return self.dialog.execute()
        
    def _isInitialized(self):
        return self.ClientId and self.AuthorizationUrl and self.TokenUrl and self.Scope

    def _setDialogStep(self, step):
        self.dialog.Title = self._getResourceString().resolveString("OAuth2Dialog.Step%s.Title" % (step,))
        self.dialog.Model.Step = step

    def _getAuthorizationFullUrl(self):
        params = {}
        params["client_id"] = self.ClientId
        params["redirect_uri"] = self.redirect
        params["response_type"] = "code"
        params["scope"] = self.Scope
        params["access_type"] = "offline"
        params["code_challenge_method"] = "S256"
        params["code_challenge"] = self._getChallengeCode()
        params["login_hint"] = self.UserName
        params["hl"] = self.locale.Language
        return requests.Request("GET", self.AuthorizationUrl, params=params).prepare().url

    def _getChallengeCode(self):
        code = hashlib.sha256(self.secret).digest()
        padding = {0:0, 1:2, 2:1}[len(code) % 3]
        challenge = base64.urlsafe_b64encode(code)
        return challenge[:len(challenge)-padding]

    def _saveSettings(self):
        if self.ClientId != self.dialog.getControl("ClientId").getText():
            self.ClientId = self.dialog.getControl("ClientId").getText()
            self.settings["ClientId"] = self._getNamedValue("ClientId", self.ClientId)
        if self.AuthorizationUrl != self.dialog.getControl("AuthorizationUrl").getText():    
            self.AuthorizationUrl = self.dialog.getControl("AuthorizationUrl").getText()
            self.settings["AuthorizationUrl"] = self._getNamedValue("AuthorizationUrl", self.AuthorizationUrl)
        if self.TokenUrl != self.dialog.getControl("TokenUrl").getText():
            self.TokenUrl = self.dialog.getControl("TokenUrl").getText()
            self.settings["TokenUrl"] = self._getNamedValue("TokenUrl", self.TokenUrl)
        if self.Scope != self.dialog.getControl("Scope").getText():
            self.Scope = self.dialog.getControl("Scope").getText()
            self.settings["Scope"] = self._getNamedValue("Scope", self.Scope)
        if self._isInitialized():
            self.dialog.getControl("AuthorizationFullUrl").setText(self._getAuthorizationFullUrl())
            return True
        return False

    def _executeShell(self, url, option=""):
        self.ctx.ServiceManager.createInstance("com.sun.star.system.SystemShellExecute").execute(url, option, 0)

    def _getTokens(self, code, timestamp):
        data = {}
        data["client_id"] = self.ClientId
        data["redirect_uri"] = self.redirect
        data["grant_type"] = "authorization_code"
        data["code"] = code
        data["code_verifier"] = self.secret
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        response = requests.post(self.TokenUrl, headers=headers, data=data, timeout=5)
        print("_getTokens: %s" % (response.json(), ))
        self.setPropertyValues(self._getNamedValuesFromJson(response.json(), timestamp))

    def _getNamedValuesFromJson(self, response, timestamp):
        result = []
        if "refresh_token" in response:
            result.append(self._getNamedValue("RefreshToken", response["refresh_token"]))
        if "access_token" in response:
            result.append(self._getNamedValue("AccessToken", response["access_token"]))
        if "expires_in" in response:
            result.append(self._getNamedValue("TimeStamp", timestamp + int(response["expires_in"])))
        return tuple(result)

    def _refreshAccessToken(self, timestamp):
        data = {}
        data["client_id"] = self.ClientId
        data["refresh_token"] = self.RefreshToken
        data["grant_type"] = "refresh_token"
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        response = requests.post(self.TokenUrl, headers=headers, data=data, timeout=5)
        print("_refreshAccessToken: %s" % (response.json(), ))
        self.setPropertyValues(self._getNamedValuesFromJson(response.json(), timestamp))
        
    def _getNamedValue(self, name, value):
        namedvalue = uno.createUnoStruct("com.sun.star.beans.NamedValue")
        namedvalue.Name = name
        namedvalue.Value = value
        return namedvalue

    def _getConfiguration (self, nodepath, update=False):
        provider = self.ctx.ServiceManager.createInstance("com.sun.star.configuration.ConfigurationProvider")
        service = "com.sun.star.configuration.ConfigurationUpdateAccess" if update else \
                  "com.sun.star.configuration.ConfigurationAccess"
        namedvalue = self._getNamedValue("nodepath", nodepath)
        return provider.createInstanceWithArguments(service, (namedvalue,))

    def _getResourceLocation(self):
        identifier = "com.gmail.prrvchr.extensions.gMailOOo"
        provider = self.ctx.getValueByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
        return "%s/gMailOOo" % (provider.getPackageLocation(identifier))

    def _getResourceString(self):
        resource = "com.sun.star.resource.StringResourceWithLocation"
        arguments = (self._getResourceLocation(), True, self.locale, "DialogStrings", "", self)
        return self.ctx.ServiceManager.createInstanceWithArgumentsAndContext(resource, arguments, self.ctx)

    def _getCurrentLocale(self):
        configuration = self._getConfiguration("/org.openoffice.Setup/L10N")
        parts = configuration.getByName("ooLocale").split("-")
        locale = uno.createUnoStruct("com.sun.star.lang.Locale")
        locale.Language = parts[0]
        if len(parts) == 2:
            locale.Country = parts[1]
        else:
            service = self.ctx.ServiceManager.createInstance("com.sun.star.i18n.LocaleData")
            locale.Country = service.getLanguageCountryInfo(locale).Country
        return locale


g_ImplementationHelper.addImplementation(PyOAuth2Service,                                            # UNO object class
                                         g_ImplementationName,                                       # Implementation name
                                        (g_ImplementationName,))                                     # List of implemented services
