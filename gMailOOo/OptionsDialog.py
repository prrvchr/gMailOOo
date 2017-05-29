#!
# -*- coding: utf_8 -*-

import uno
import unohelper

#interfaces
from com.sun.star.lang import XServiceInfo
from com.sun.star.awt import XContainerWindowEventHandler

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = "com.gmail.prrvchr.extensions.gMailOOo.OptionsDialog"

g_SettingNodePath = "com.gmail.prrvchr.extensions.gMailOOo/MailMergeWizard"
g_ConnectionTimeout = "ConnectionTimeout"
g_ConnectionSecurity = "ConnectionSecurity"
g_AuthenticationMethod = "AuthenticationMethod"

# main class
class PyOptionsDialog(unohelper.Base, XServiceInfo, XContainerWindowEventHandler):
    def __init__(self, ctx):
        self.ctx = ctx
        self.dialog = None
        return

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
                self._setUnsecure()
                return True
            elif method == "Secure":
                self._setSecure()
                return True
            elif method == "OAuth2":
                self._setOAuth2()
                return True
            elif method == "OAuth2Setup":
                self._setOAuth2Setup()
                return True
        return False
    def getSupportedMethodNames(self):
        return ("external_event", "Unsecure", "Secure", "OAuth2", "OAuth2Setup")

    # XServiceInfo
    def supportsService(self, serviceName):
        return g_ImplementationHelper.supportsService(g_ImplementationName, serviceName)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    def _setUnsecure(self):
        if self.dialog.getControl("OptionButton6").getState():
            self.dialog.getControl("OptionButton5").setState(True)
        self.dialog.getControl("OptionButton6").Model.Enabled = False

    def _setSecure(self):
        self.dialog.getControl("OptionButton6").Model.Enabled = True

    def _setOAuth2(self):
        timestamp = self._getConfigSetting(g_SettingNodePath, "ExpiresTimeStamp")
        if timestamp == 0:
            self._setOAuth2Setup()

    def _setOAuth2Setup(self):
        self.dialog.getControl("NumericField1").setValue(0)
        component = self.ctx.ServiceManager.createInstanceWithArgumentsAndContext("com.gmail.prrvchr.extensions.gMailOOo.OAuth2Service", (), self.ctx)
#        component = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.OAuth2Service", self.ctx)
        self.dialog.getControl("NumericField1").setValue(1)
        timestamp = component.getByName("ExpiresTimeStamp")
        self.dialog.getControl("NumericField1").setValue(2)
#        self.dialog.getControl("NumericField1").setValue(timestamp)
#        component = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.OAuth2Service", self.ctx)
#        arg = ("docmd", "arg1")
#        try:
#        component.invoke("getAuthenticationCmd", (), (), ())
#        except:
#        component.initialize(())

    def _loadSetting(self):
        self.dialog.getControl("NumericField1").setValue(self._getConfigSetting(g_SettingNodePath, g_ConnectionTimeout))
        security = self._getConfigSetting(g_SettingNodePath, g_ConnectionSecurity) + 1
        self.dialog.getControl("OptionButton%s" % (security)).setState(True)
#        self.dialogcomponent = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.OAuth2Service", self.ctx).getControl("OptionButton%s" % (security)).setState(True)
        authentication = self._getConfigSetting(g_SettingNodePath, g_AuthenticationMethod) + 4
        self.dialog.getControl("OptionButton%s" % (authentication)).setState(True)
        return

    def _saveSetting(self):
        self._setConfigSetting(g_SettingNodePath, g_ConnectionTimeout, int(self.dialog.getControl("NumericField1").getValue()))
        for i in range(1, 4):
            if self.dialog.getControl("OptionButton%s" % (i)).getState():
                self._setConfigSetting(g_SettingNodePath, g_ConnectionSecurity, i - 1)
        for i in range(4, 7):
            if self.dialog.getControl("OptionButton%s" % (i)).getState():
                self._setConfigSetting(g_SettingNodePath, g_AuthenticationMethod, i - 4)
        return

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
        return

# uno implementation
g_ImplementationHelper.addImplementation( \
        PyOptionsDialog,                                                       # UNO object class
        g_ImplementationName,                                                  # Implementation name
        ("com.sun.star.lang.XServiceInfo",
        "com.sun.star.awt.XContainerWindowEventHandler"), )                    # List of implemented services
