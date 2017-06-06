#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo
from com.sun.star.lang import XInitialization
from com.sun.star.awt import XDialogEventHandler
from com.sun.star.util import XStringEscape

import urllib.parse, urllib.request
import json
import time
import base64

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = "com.gmail.prrvchr.extensions.gMailOOo.OAuth2Service"

g_SettingNodePath = "com.gmail.prrvchr.extensions.gMailOOo/MailMergeWizard"

g_ClientId = "429153119813-sg5ks5gfu737nm2iitm5v1um3kr0a1fo.apps.googleusercontent.com"
g_ClientSecret = "1VPCRFBpgYZGvTwoyMFsx6Or"
# The URL root for accessing Google Accounts.
GOOGLE_ACCOUNTS_BASE_URL = "https://accounts.google.com"
# Hardcoded dummy redirect URI for non-web apps.
REDIRECT_URI = "urn:ietf:wg:oauth:2.0:oob"


class PyOAuth2Service(unohelper.Base, XServiceInfo, XInitialization, XDialogEventHandler, XStringEscape):
    def __init__(self, ctx):
        self.ctx = ctx
        self.dialog = None

    # XInitialization
    def initialize(self, *arg):
        url = self._getPermissionUrl()
        self._openDialog(1, url)

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        if self.dialog == dialog:
            step = dialog.Model.Step
            if method == "DialogOk":
                if step == 1:
                    self._openBrowser()
                    dialog.Model.Step = 2
                elif step == 2:
                    authorization = dialog.getControl("TextField2").getText()
                    if authorization != "":
                        self._getTokens(authorization)
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
    def supportsService(self, serviceName):
        return g_ImplementationHelper.supportsService(g_ImplementationName, serviceName)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    def _openBrowser(self):
        url = self._getPermissionUrl()
        shell = self.ctx.ServiceManager.createInstance("com.sun.star.system.SystemShellExecute")
        shell.execute(url, "", 0)

    def _openDialog(self, step, text=""):
        provider = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.DialogProvider2", self.ctx)
        self.dialog = provider.createDialogWithHandler("vnd.sun.star.script:gMailOOo.Dialog?location=application", self)
        self.dialog.Model.Step = step
        if text != "":
            self.dialog.getControl("TextField%s" % (step)).setText(text)
        self.dialog.execute()
        self.dialog.dispose()
        self.dialog = None

    def _getPermissionUrl(self, scope="https://www.googleapis.com/auth/gmail.send"):
        params = {}
        params["client_id"] = g_ClientId
        params["redirect_uri"] = REDIRECT_URI
        params["scope"] = scope
        params["prompt"] = "consent"
        params["response_type"] = "code"
        return "%s?%s" % (self._getAccountsUrl("o/oauth2/auth"), urllib.parse.urlencode(params))

    def _getAccountsUrl(self, command):
        return "%s/%s" % (GOOGLE_ACCOUNTS_BASE_URL, command)

    def _getTokens(self, authorization):
        timestamp = int(time.time())
        response = self._generateTokens(authorization)
        self._saveResponse(response, timestamp)

    def _generateTokens(self, authorization):
        params = {}
        params["client_id"] = g_ClientId
        params["client_secret"] = g_ClientSecret
        params["code"] = authorization
        params["grant_type"] = "authorization_code"
        params["redirect_uri"] = REDIRECT_URI
        request_url = self._getAccountsUrl("o/oauth2/token")
        request_param = urllib.parse.urlencode(params).encode("utf-8")
        request = urllib.request.Request(request_url, request_param)
        response = urllib.request.urlopen(request, timeout=5).read()
        return json.loads(response.decode("utf-8"))

    def _isAccessTokenExpired(self, timestamp):
        expirestimestamp = self._getConfigSetting(g_SettingNodePath, "ExpiresTimeStamp")
        if expirestimestamp <= timestamp:
            return True
        return False

    def _getRefreshToken(self):
        self._openBrowser()
        self._openDialog(2)
        refreshtoken = self._getConfigSetting(g_SettingNodePath, "RefreshToken")
        return refreshtoken

    def _refreshAccessToken(self, timestamp):
        refreshtoken = self._getConfigSetting(g_SettingNodePath, "RefreshToken")
        if refreshtoken == "":
            refreshtoken = self._getRefreshToken()
        response = self._generateAccessToken(refreshtoken)
        self._saveResponse(response, timestamp)

    def _generateAccessToken(self, refreshtoken):
        params = {}
        params["client_id"] = g_ClientId
        params["client_secret"] = g_ClientSecret
        params["refresh_token"] = refreshtoken
        params["grant_type"] = "refresh_token"
        request_url = self._getAccountsUrl("o/oauth2/token")
        request_param = urllib.parse.urlencode(params).encode("utf-8")
        request = urllib.request.Request(request_url, request_param)
        response = urllib.request.urlopen(request, timeout=5).read()
        return json.loads(response.decode("utf-8"))

    def _getAuthenticationString(self, username, encode=False):
        timestamp = int(time.time())
        if self._isAccessTokenExpired(timestamp):
            self._refreshAccessToken(timestamp)
        accesstoken = self._getConfigSetting(g_SettingNodePath, "AccessToken")
        authstring = "user=%s\1auth=Bearer %s\1\1" % (username, accesstoken)
        if encode:
          authstring = base64.b64encode(authstring.encode("ascii"))
        return authstring
        
    def _saveResponse(self, response, timestamp):
        if "refresh_token" in response:
            self._setConfigSetting(g_SettingNodePath, "RefreshToken", response["refresh_token"])
        self._setConfigSetting(g_SettingNodePath, "AccessToken", response["access_token"])
        self._setConfigSetting(g_SettingNodePath, "ExpiresTimeStamp", timestamp + int(response["expires_in"]))

    def _getPropertyValue(self, nodepath):
        args = []
        arg = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        arg.Name = "nodepath"
        arg.Value = nodepath
        args.append(arg)
        return tuple(args)

    def _getConfigSetting(self, nodepath, property):
        config = self.ctx.ServiceManager.createInstance("com.sun.star.configuration.ConfigurationProvider")
        access = config.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", self._getPropertyValue(nodepath))
        return access.getByName(property)

    def _setConfigSetting(self, nodepath, property, value):
        config = self.ctx.ServiceManager.createInstance("com.sun.star.configuration.ConfigurationProvider")
        access = config.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", self._getPropertyValue(nodepath))
        access.replaceByName(property, value)
        access.commitChanges()


g_ImplementationHelper.addImplementation( \
        PyOAuth2Service,                                                       # UNO object class
        g_ImplementationName,                                                  # Implementation name
        ("com.sun.star.lang.XServiceInfo",
        "com.sun.star.lang.XInitialization",
        "com.sun.star.awt.XDialogEventHandler",
        "com.sun.star.util.XStringEscape"), )                                  # List of implemented services
