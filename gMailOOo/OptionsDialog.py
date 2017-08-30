#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo
from com.sun.star.awt import XContainerWindowEventHandler
from com.sun.star.beans import PropertyValue

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = "com.gmail.prrvchr.extensions.gMailOOo.OptionsDialog"


class PyOptionsDialog(unohelper.Base, XServiceInfo, XContainerWindowEventHandler):
    def __init__(self, ctx):
        self.ctx = ctx
        self.dialog = None
        self.configuration = "com.gmail.prrvchr.extensions.gMailOOo/MailMergeWizard"
        return

    # XContainerWindowEventHandler
    def callHandlerMethod(self, dialog, event, method):
        if dialog.Model.Name == "OptionsDialog":
            if method == "external_event":
                if event == "ok":
                    configuration = self._getConfiguration(self.configuration, True)
                    configuration.replaceByName("ConnectionTimeout", int(self.dialog.getControl("NumericField1").getValue()))
                    for i in range(1, 4):
                        if self.dialog.getControl("OptionButton%s" % (i)).getState():
                            configuration.replaceByName("ConnectionSecurity", i - 1)
                    for i in range(4, 7):
                        if self.dialog.getControl("OptionButton%s" % (i)).getState():
                            configuration.replaceByName("AuthenticationMethod", i - 4)
                    configuration.commitChanges()
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
                if self.dialog.getControl("OptionButton6").getState():
                    self.dialog.getControl("OptionButton5").setState(True)
                self.dialog.getControl("OptionButton6").Model.Enabled = False
                self.dialog.getControl("CommandButton1").Model.Enabled = False
                self.dialog.getControl("CommandButton2").Model.Enabled = False
                return True
            elif method == "Secure":
                self.dialog.getControl("OptionButton6").Model.Enabled = True
                self.dialog.getControl("CommandButton1").Model.Enabled = True
                self.dialog.getControl("CommandButton2").Model.Enabled = True
                return True
            elif method == "OAuth2":
                timestamp = self._getConfiguration(self.configuration).getByName("ExpiresTimeStamp")
                if timestamp == 0:
                    self._setOAuth2Setup()
                return True
            elif method == "OAuth2Setup":
                self._setOAuth2Setup()
                return True
        return False
    def getSupportedMethodNames(self):
        return ("external_event", "Unsecure", "Secure", "OAuth2", "OAuth2Setup")

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    def _setOAuth2Setup(self):
        component = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.OAuth2Service", self.ctx)
        component.trigger(2)

    def _loadSetting(self):
        self.dialog.getControl("NumericField1").setValue(self._getConfiguration(self.configuration).getByName("ConnectionTimeout"))
        security = self._getConfiguration(self.configuration).getByName("ConnectionSecurity") + 1
        self.dialog.getControl("OptionButton%s" % (security)).setState(True)
        authentication = self._getConfiguration(self.configuration).getByName("AuthenticationMethod") + 4
        self.dialog.getControl("OptionButton%s" % (authentication)).setState(True)

    def _getConfiguration(self, nodepath, update=False):
        value = uno.Enum("com.sun.star.beans.PropertyState", "DIRECT_VALUE")
        config = self.ctx.ServiceManager.createInstance("com.sun.star.configuration.ConfigurationProvider")
        service = "com.sun.star.configuration.ConfigurationUpdateAccess" if update else "com.sun.star.configuration.ConfigurationAccess"
        return config.createInstanceWithArguments(service, (PropertyValue("nodepath", -1, nodepath, value), ))


g_ImplementationHelper.addImplementation(PyOptionsDialog,                                            # UNO object class
                                         g_ImplementationName,                                       # Implementation name
                                         ("com.sun.star.lang.XServiceInfo",
                                         "com.sun.star.awt.XContainerWindowEventHandler"), )         # List of implemented services
