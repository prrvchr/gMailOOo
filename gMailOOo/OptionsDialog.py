#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo
from com.sun.star.awt import XContainerWindowEventHandler
from com.sun.star.mail import XAuthenticator
from com.sun.star.util import XChangesListener
from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.mail.MailServiceType import POP3
from com.sun.star.mail.MailServiceType import IMAP

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = "com.gmail.prrvchr.extensions.gMailOOo.OptionsDialog"


class PyOptionsDialog(unohelper.Base, XServiceInfo, XContainerWindowEventHandler, XChangesListener):
    def __init__(self, ctx):
        self.ctx = ctx
        self.dialog = None
        self.nodepath = "com.gmail.prrvchr.extensions.gMailOOo/MailMergeWizard"
        self.configuration = self._getConfiguration("/org.openoffice.Office.Writer/MailMergeWizard", True)
        self.elementschange = ()

    # XChangesListener
    def changesOccurred(self, changesevent):
        self.elementschange += changesevent.Changes
        print("PyOptionsDialog.changesOccurred: %s \n\n %s" % (changesevent.Base, changesevent.Changes))

    # XContainerWindowEventHandler
    def callHandlerMethod(self, dialog, event, method):
        if dialog.Model.Name == "OptionsDialog":
            if method == "external_event":
                if event == "ok":
                    self._saveSetting()
                    return True
                elif event == "back":
                    self._loadSetting()
                    return True
                elif event == "initialize":
                    if self.dialog is None:
                        self.dialog = dialog
                    self._loadSetting()
                    return True
            elif method == "Unsecure":
                self._setSecureState(False)
                return True
            elif method == "Secure":
                self._setSecureState(True)
                return True
            elif method == "Connect":
                self._initConnection()
                return True
        return False
    def getSupportedMethodNames(self):
        return ("external_event", "Unsecure", "Secure", "Connect")

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    def _testConnection(self):
        try:
            print("_initOAuth2 1")
            mailservicetype = self._getCurrentMailServiceType()
            provider = self._getMailServiceProvider(mailservicetype)
            connectioncontext = self.ctx.ServiceManager.createInstanceWithArgumentsAndContext("com.gmail.prrvchr.extensions.gMailOOo.ConnectionContext", (self._getNamedValue("MailServiceType", mailservicetype),), self.ctx)
            connectioncontext.getPropertyValue("Configuration").addChangesListener(self)
            print("_initOAuth2 2")
            authenticator = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.Authenticator", self.ctx)
            authenticator.getPropertyValue("Configuration").addChangesListener(self)
            provider.connect(connectioncontext, authenticator)
            print("_initOAuth2 3 isConnected %s" % (provider.isConnected,))
            if mailservicetype != SMTP:
                outprovider = self._getMailServiceProvider(SMTP)
                connectioncontext.setPropertyValue("MailServiceType", SMTP)
                outprovider.connect(connectioncontext, authenticator)
                print("_initOAuth2 4 isConnected %s" % (outprovider.isConnected,))
            connectioncontext.getPropertyValue("Configuration").removeChangesListener(self)
            authenticator.getPropertyValue("Configuration").removeChangesListener(self)
        except Exception as e:
            print("_initOAuth2.Error: %s" % (e,))

    def _setSecureState(self, state):
        if not state and self.dialog.getControl("OptionButton6").getState():
            self.dialog.getControl("OptionButton5").setState(True)
        self.dialog.getControl("OptionButton6").Model.Enabled = state

    def _loadSetting(self):
        configuration = self._getConfiguration(self.nodepath)
        self.dialog.getControl("NumericField1").setValue(configuration.getByName("ConnectionTimeout"))
        button = 1 if not self.configuration.getByName("IsSecureConnection") else \
                (2 if not configuration.getByName("IsSecureLevel2") else \
                 3)
        self.dialog.getControl("OptionButton%s" % (button)).setState(True)
        button = 4 if not self.configuration.getByName("IsAuthentication") else \
                (5 if not configuration.getByName("IsOAuth2") else \
                 6)
        self.dialog.getControl("OptionButton%s" % (button)).setState(True)
        self._setSecureState(self._isOAuth2Supported() )

    def _getCurrentMailServiceType(self):
        return SMTP if not self.configuration.getByName("IsSMPTAfterPOP") else \
              (POP3 if self.configuration.getByName("InServerIsPOP") else \
               IMAP)

    def _getMailServiceProvider(self, mailservicetype=None):
        if mailservicetype is None:
            mailservicetype = self._getCurrentMailServiceType()
        provider = self.ctx.ServiceManager.createInstanceWithContext("org.openoffice.pyuno.MailServiceProvider", self.ctx)
        return provider.create(mailservicetype)

    def _isOAuth2Supported(self):
        return "OAuth2" in self._getMailServiceProvider().getSupportedConnectionTypes()

    def _saveSetting(self):
        try:
            for elementchange in self.elementschange:
                print("PyOptionsDialog._saveSetting: %s %s" % (elementchange.Accessor, elementchange.Element))
                self.configuration.replaceByName(elementchange.Accessor, elementchange.Element)
            configuration = self._getConfiguration(self.nodepath, True)
            configuration.replaceByName("ConnectionTimeout", int(self.dialog.getControl("NumericField1").getValue()))
            self.configuration.replaceByName("IsSecureConnection", self.dialog.getControl("OptionButton1").getState() == 0)
            configuration.replaceByName("IsSecureLevel2", self.dialog.getControl("OptionButton3").getState() == 1)
            self.configuration.replaceByName("IsAuthentication", self.dialog.getControl("OptionButton4").getState() == 0)
            configuration.replaceByName("IsOAuth2", self.dialog.getControl("OptionButton6").getState() == 1)
            if self.configuration.hasPendingChanges():
                self.configuration.commitChanges()
            if configuration.hasPendingChanges():
                configuration.commitChanges()
        except Exception as e:
            print("PyOptionsDialog._saveSetting: Error;: %s" % (e,))

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


g_ImplementationHelper.addImplementation(PyOptionsDialog,                                            # UNO object class
                                         g_ImplementationName,                                       # Implementation name
                                        (g_ImplementationName,))                                     # List of implemented services
