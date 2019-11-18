#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo, XInitialization, XLocalizable
from com.sun.star.uno import XCurrentContext
from com.sun.star.task import XInteractionHandler
from com.sun.star.beans import XPropertySet
from com.sun.star.mail.MailServiceType import SMTP

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = "com.gmail.prrvchr.extensions.gMailOOo.ConnectionContext"


class PyConnectionContext(unohelper.Base, XServiceInfo, XInitialization, XLocalizable, XCurrentContext, XInteractionHandler, XPropertySet):
    def __init__(self, ctx, *namedvalues):
        self.ctx = ctx
        self.dialog = None
        self.dialogurl = "vnd.sun.star.script:gMailOOo.ConnectionContextDialog?location=application"
        self.locale = self._getCurrentLocale()
        self.nodepath = "com.gmail.prrvchr.extensions.gMailOOo/MailMergeWizard"
        self.MailServiceType = None
        self.Configuration = self._getConfiguration("/org.openoffice.Office.Writer/MailMergeWizard", True)
        self.initialize(namedvalues)

    # XInitialization
    def initialize(self, namedvalues=()):
        for namedvalue in namedvalues:
            if hasattr(namedvalue, "Name") and hasattr(namedvalue, "Value"):
                self.setPropertyValue(namedvalue.Name, namedvalue.Value)

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

    # XCurrentContext
    def getValueByName(self, valuename):
        if valuename == "ServerName":
            name = "MailServer" if self.MailServiceType == SMTP else \
                   "InServerName"
            servername = self.Configuration.getByName(name)
            if not servername:
                if self._openDialog(1):
                    servername = self.dialog.getControl("TextField1").getText()
                    if servername:
                        self.Configuration.replaceByName(name, servername)
                        self.Configuration.commitChanges()
                self.dialog.dispose()
                self.dialog = None
            return servername
        elif valuename == "Port":
            name = "MailPort" if self.MailServiceType == SMTP else \
                   "InServerPort"
            port = self.Configuration.getByName(name)
            if not port:
                if self._openDialog(2):
                    port = self.dialog.getControl("TextField2").getText()
                    if port:
                        self.Configuration.replaceByName(name, port)
                        self.Configuration.commitChanges()
                self.dialog.dispose()
                self.dialog = None
            return port
        elif valuename == "ConnectionTimeout":
            configuration = self._getConfiguration(self.nodepath)
            return configuration.getByName("ConnectionTimeout")
        elif valuename == "ConnectionType":
            configuration = self._getConfiguration(self.nodepath)
            return "Insecure" if not self.Configuration.getByName("IsSecureConnection") else \
                  ("Ssl" if not configuration.getByName("IsSecureLevel2") else \
                   "Tls")
        elif valuename == "AuthenticationType":
            configuration = self._getConfiguration(self.nodepath)
            if self.MailServiceType == SMTP and self.Configuration.getByName("IsSMPTAfterPOP"):
                return "None" 
            else:
                return "None" if not self.Configuration.getByName("IsAuthentication") else \
                      ("Login" if not configuration.getByName("IsOAuth2") else \
                       "OAuth2")

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    def _openDialog(self, step):
        provider = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.DialogProvider", self.ctx)
        self.dialog = provider.createDialog(self.dialogurl)
        self.dialog.Title = self._getResourceString().resolveString("ConnectionContextDialog.Step%s.Title" % (step,))
        self.dialog.Model.Step = step
        return self.dialog.execute()

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


g_ImplementationHelper.addImplementation(PyConnectionContext,                                        # UNO object class
                                         g_ImplementationName,                                       # Implementation name
                                        (g_ImplementationName,))                                     # List of implemented services
