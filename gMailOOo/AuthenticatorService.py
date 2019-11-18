#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo, XInitialization, XLocalizable
from com.sun.star.mail import XAuthenticator
from com.sun.star.util import XStringEscape
from com.sun.star.task import XInteractionHandler
from com.sun.star.beans import XPropertySet

import base64

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = "com.gmail.prrvchr.extensions.gMailOOo.Authenticator"


class PyAuthenticator(unohelper.Base, XServiceInfo, XInitialization, XLocalizable, XAuthenticator, XStringEscape, XInteractionHandler, XPropertySet):
    def __init__(self, ctx, *namedvalues):
        self.ctx = ctx
        self.dialog = None
        self.dialogurl = "vnd.sun.star.script:gMailOOo.AuthenticatorDialog?location=application"
        self.locale = self._getCurrentLocale()
        self.nodepath = "com.gmail.prrvchr.extensions.gMailOOo/MailMergeWizard"
        self.Configuration = self._getConfiguration("/org.openoffice.Office.Writer/MailMergeWizard", True)
        self.initialize(namedvalues)

    # XInitialization
    def initialize(self, namedvalues=()):
        for namedvalue in namedvalues:
            if hasattr(namedvalue, "Name") and hasattr(namedvalue, "Value") and hasattr(self, namedvalue.Name):
                setattr(self, namedvalue.Name, namedvalue.Value)

    # XInteractionHandler
    def handle(self, requester):
        pass

    # XLocalizable
    def setLocale(self, locale):
        self.locale = locale
    def getLocale(self):
        return self.locale

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

    # XAuthenticator
    def getUserName(self):
        name = "InServerUserName" if self.Configuration.getByName("IsSMPTAfterPOP") else \
               "MailUserName"
        username = self.Configuration.getByName(name)
        if not username:
            if self._openDialog(1):
                username = self.dialog.getControl("TextField1").getText()
                if username:
                    self.Configuration.replaceByName(name, username)
                    self.Configuration.commitChanges()
            self.dialog.dispose()
            self.dialog = None
        return username
    def getPassword(self):
        name = "InServerPassword" if self.Configuration.getByName("IsSMPTAfterPOP") else \
               "MailPassword"
        password = self.Configuration.getByName(name)
        if not password:
            if self._openDialog(2):
                password = self.dialog.getControl("TextField2").getText()
#                if password:
#                    self.Configuration.replaceByName(name, password)
#                    self.Configuration.commitChanges()
            self.dialog.dispose()
            self.dialog = None
        return password

    # XStringEscape
    def escapeString(self, mailserver):
        authstring = base64.b64encode(self.unescapeString(mailserver).encode("ascii"))
        return authstring
    def unescapeString(self, mailserver):
        mailaddress = self._getUserMail()
        accesstoken = self._getOAuth2Tokens(mailaddress, mailserver)
        authstring = "user=%s\1auth=Bearer %s\1\1" % (mailaddress, accesstoken)
        return authstring

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    def _getUserMail(self):
        mailaddress = self.Configuration.getByName("MailAddress")
        if not mailaddress:
            if self._openDialog(3):
                mailaddress = self.dialog.getControl("TextField3").getText()
                if mailaddress:
                    self.Configuration.replaceByName("MailAddress", mailaddress)
                    self.Configuration.commitChanges()
            self.dialog.dispose()
            self.dialog = None
        return mailaddress

    def _getOAuth2Tokens(self, mailaddress, mailserver):
        configuration = self._getConfiguration(self.nodepath, True)
        oauth2servers = configuration.getByName("OAuth2servers")
        oauth2server = oauth2servers.getByName(mailserver) if oauth2servers.hasByName(mailserver) else \
                       oauth2servers.getByName("default")
        settings, tokens = self._getNamedValuesFromConfiguration(oauth2server, mailaddress)
        oauth2service = self.ctx.ServiceManager.createInstanceWithArgumentsAndContext("com.gmail.prrvchr.extensions.gMailOOo.OAuth2Service", settings, self.ctx)
        namedvalues, dummy, dummy1 = oauth2service.invoke(tokens, (0,), ())
        if len(namedvalues) and len(mailserver):
            if not oauth2servers.hasByName(mailserver):
                oauth2servers.insertByName(mailserver, oauth2servers.createInstance())
                oauth2server = oauth2servers.getByName(mailserver)
            self._setConfigurationFromNamedValues(oauth2server, namedvalues)
            configuration.commitChanges()
        return oauth2server.getByName("AccessToken")

    def _getNamedValuesFromConfiguration(self, configuration, mailaddress):
        settings = (self._getNamedValue("ClientId", configuration.getByName("ClientId")),
                    self._getNamedValue("AuthorizationUrl", configuration.getByName("AuthorizationUrl")),
                    self._getNamedValue("TokenUrl", configuration.getByName("TokenUrl")),
                    self._getNamedValue("Scope", configuration.getByName("Scope")),
                    self._getNamedValue("UserName", mailaddress))
        tokens = (self._getNamedValue("RefreshToken", configuration.getByName("RefreshToken")),
                  self._getNamedValue("AccessToken", configuration.getByName("AccessToken")),
                  self._getNamedValue("TimeStamp", configuration.getByName("TimeStamp")))
        return settings, tokens

    def _setConfigurationFromNamedValues(self, configuration, namedvalues):
        for namedvalue in namedvalues:
            if hasattr(namedvalue, "Name") and hasattr(namedvalue, "Value"):
                if configuration.hasByName(namedvalue.Name):
                    configuration.replaceByName(namedvalue.Name, namedvalue.Value)
                else:
                    configuration.insertByName(namedvalue.Name, namedvalue.Value)

    def _openDialog(self, step):
        provider = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.DialogProvider", self.ctx)
        self.dialog = provider.createDialog(self.dialogurl)
        self.dialog.Title = self._getResourceString().resolveString("AuthenticatorDialog.Step%s.Title" % (step,))
        self.dialog.Model.Step = step
        return self.dialog.execute()

    def _getNamedValue(self, name, value):
        namedvalue = uno.createUnoStruct("com.sun.star.beans.NamedValue")
        namedvalue.Name = name
        namedvalue.Value = value
        return namedvalue

    def _getConfiguration(self, nodepath, update=False):
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


g_ImplementationHelper.addImplementation(PyAuthenticator,                                            # UNO object class
                                         g_ImplementationName,                                       # Implementation name
                                        (g_ImplementationName,))                                     # List of implemented services
