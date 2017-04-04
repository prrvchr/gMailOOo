#!
# -*- coding: utf_8 -*-

import uno
import unohelper

#interfaces
from com.sun.star.lang import XServiceInfo
from com.sun.star.awt import XActionListener
from com.sun.star.util import XChangesListener
from com.sun.star.awt import XContainerWindowEventHandler

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = "com.gmail.prrvchr.extensions.gMailOOo.OptionsDialog"

g_SettingNodePath = "com.gmail.prrvchr.extensions.gMailOOo/MailMergeWizard"
g_ConnectionTimeout = "ConnectionTimeout"
g_ConnectionSecurity = "ConnectionSecurity"
g_AuthenticationMethod = "AuthenticationMethod"

# main class
class PyOptionsDialog(unohelper.Base, XServiceInfo, XActionListener, XChangesListener, XContainerWindowEventHandler):
    def __init__(self, ctx):
        self.ctx = ctx
        self.dialog = None
        return

    # XChangesListener
    def changesOccurred(self, event):
        pass

    # XActionListener
    def actionPerformed(self, event):
        if self.dialog != event.Source.getContext():
            return
        if event.Source.Model.Name == "OptionButton1":
            if event.Source.getState():
                self.dialog.getControl("OptionButton6").Model.Enable = False
                if self.dialog.getControl("OptionButton6").getState():
                    self.dialog.getControl("OptionButton5").setState(True)
            else:
                self.dialog.getControl("OptionButton6").Model.Enable = True
        elif event.Source.Model.Name == "CommandButton1":
            component = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.OAuth2Service", self.ctx)
            component.initialize(())
    def disposing(self, event):
        pass

    # XContainerWindowEventHandler
    def callHandlerMethod(self, dialog, event, method):
        if dialog.Model.Name == "OptionsDialog":
            if method == "external_event":
                if event == "ok":
                    self._eventOk()
                    return True
                elif event == "back":
                    self._eventBack()
                    return True
                elif event == "initialize":
                    if self.dialog is None:
                        self.dialog = dialog
                    self._eventInitialize()
                    return True
        return False
    def getSupportedMethodNames(self):
        return ("external_event", )

    # XServiceInfo
    def supportsService(self, serviceName):
        return g_ImplementationHelper.supportsService(g_ImplementationName, serviceName)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    def _eventOk(self):
#        access = self._getConfigAccess("org.openoffice.Office.Writer/MailMergeWizard")
#        access.removeChangesListener(self)
        self._setConfigSetting(g_SettingNodePath, g_ConnectionTimeout, int(self.dialog.getControl("NumericField1").getValue()))
        for i in range(1, 4):
            if self.dialog.getControl("OptionButton%s" % (i)).getState():
                self._setConfigSetting(g_SettingNodePath, g_ConnectionSecurity, i - 1)
        for i in range(4, 7):
            if self.dialog.getControl("OptionButton%s" % (i)).getState():
                self._setConfigSetting(g_SettingNodePath, g_AuthenticationMethod, i - 4)
        return

    def _eventBack(self):
        self.dialog.getControl("NumericField1").setValue(self._getConfigSetting(g_SettingNodePath, g_ConnectionTimeout))
        security = self._getConfigSetting(g_SettingNodePath, g_ConnectionSecurity) + 1
        self.dialog.getControl("OptionButton%s" % (security)).setState(True)
        authentication = self._getConfigSetting(g_SettingNodePath, g_AuthenticationMethod) + 4
        self.dialog.getControl("OptionButton%s" % (authentication)).setState(True)
        return

    def _eventInitialize(self):
#        access = self._getConfigAccess("org.openoffice.Office.Writer/MailMergeWizard", True)
#        access.addChangesListener(self)
        self.dialog.getControl("CommandButton1").addActionListener(self)
        self._eventBack()
        return

    def _getPropertyValue(self, nodepath):
        args = []
        arg = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        arg.Name = "nodepath"
        arg.Value = nodepath
        args.append(arg)
        return tuple(args)

    def _getConfigAccess(self, nodepath, update=False):
        config = self.ctx.ServiceManager.createInstance("com.sun.star.configuration.ConfigurationProvider")
        if update:
            service = "com.sun.star.configuration.ConfigurationAccess"
        else:
            service = "com.sun.star.configuration.ConfigurationUpdateAccess"
        return config.createInstanceWithArguments(service, self._getPropertyValue(nodepath))

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
        "com.sun.star.awt.XActionListener",
        "com.sun.star.util.XChangesListener",
        "com.sun.star.awt.XContainerWindowEventHandler"), )                    # List of implemented services
