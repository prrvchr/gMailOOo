#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo
from com.sun.star.awt import XDialogEventHandler
from com.sun.star.util import XCloseable
from com.sun.star.task import XJob

import os
import stat

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = "com.gmail.prrvchr.extensions.gMailOOo.MailmergeUpdate"
g_MailMergePython = "mailmerge.py"

class PyMailmergeUpdate(unohelper.Base, XServiceInfo, XDialogEventHandler, XCloseable, XJob):
    def __init__(self, ctx):
        self.ctx = ctx
        self.source = "%s/%s" % (self._getPackageLocation(), g_MailMergePython)
        self.target = "%s/%s" % (self._getPathSubstitution("$(prog)"), g_MailMergePython)
        self.dialog = None
        self.restart = False

    # XJob
    def execute(self, *args):
        if self._isMailServiceNeedUpdate():
            self._openDialog()

    # XCloseable
    def close(self, DeliverOwnership):
        if self.dialog is not None:
            self.dialog.endExecute()

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        if self.dialog == dialog:
            step = dialog.Model.Step
            if method == "DialogOk":
                if step == 3:
                    self._updateMailService()
                elif step == 4:
                     dialog.endExecute()
                elif step == 5:
                    self.restart = True
                    dialog.endExecute()
                return True
            elif method == "DialogCancel":
                dialog.endExecute()
                return True
        return False
    def getSupportedMethodNames(self):
        return ("DialogOk", "DialogCancel")

    # XServiceInfo
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)
    def supportsService(self, ServiceName):
        return g_ImplementationHelper.supportsService(g_ImplementationName, ServiceName)

    def _openDialog(self):
        provider = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.DialogProvider2", self.ctx)
        self.dialog = provider.createDialogWithHandler("vnd.sun.star.script:gMailOOo.Dialog?location=application", self)
        self.dialog.getControl("TextField3").Text = uno.fileUrlToSystemPath(self.target)
        self.dialog.Model.Step = 3
        self.dialog.execute()
        self.dialog.dispose()
        self.dialog = None
        if self.restart:
            desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
            desktop.terminate()
 
    def _isMailServiceNeedUpdate(self):
        service = self.ctx.ServiceManager.createInstance("com.sun.star.ucb.SimpleFileAccess")
        if service.exists(self.target):
            if service.getSize(self.target) == service.getSize(self.source):
                return False
        return True

    def _getScriptPath(self, script):
        return uno.fileUrlToSystemPath("%s/%s" % (self._getPackageLocation(), script))

    def _getPathSubstitution(self, path):
        pathsubstitution = self.ctx.ServiceManager.createInstance("com.sun.star.util.PathSubstitution")
        return pathsubstitution.getSubstituteVariableValue(path)

    def _getPackageLocation(self):
        identifier = "com.gmail.prrvchr.extensions.gMailOOo"
        pip = self.ctx.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
        return pip.getPackageLocation(identifier)

    def _updateMailService(self):
        command, text = self._getUpdateCommand()
        status = os.system(command)
        if status or self._isMailServiceNeedUpdate():
            self.dialog.getControl("TextField4").Text = text
            self.dialog.Model.Step = 4
        else:
            self.dialog.Model.Step = 5

    def _getUpdateCommand(self):
        source = uno.fileUrlToSystemPath(self.source)
        target = uno.fileUrlToSystemPath(self.target)
        if os.name == "nt":
            script = self._getScriptPath("MailmergeUpdate.ps1")
            command = "powershell.exe -ExecutionPolicy ByPass -File \"%s\" -source \"%s\" -target \"%s\"" % (script, source, target)
            text = "copy \"%s\" \"%s\"" % (source, target)
        else:
            script = self._getScriptPath("MailmergeUpdate.sh")
            mode = os.stat(script).st_mode
            executable = mode & stat.S_IXUSR
            if not executable:
                os.chmod(script, (mode | stat.S_IXUSR))
            command = "x-terminal-emulator -e '%s' --source '%s' --target '%s'" % (script, source, target)
            text = "sudo cp '%s' '%s'" % (source, target)
        return (command, text)


g_ImplementationHelper.addImplementation( \
        PyMailmergeUpdate,                                                     # UNO object class
        g_ImplementationName,                                                  # Implementation name
        ("com.sun.star.lang.XServiceInfo",
        "com.sun.star.awt.XDialogEventHandler",
        "com.sun.star.util.XCloseable",
        "com.sun.star.task.XJob",),)                                           # List of implemented services
