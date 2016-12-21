#   Copyright (C) 2016 Canopus GeoInformatics Ltd.
#
#    This program is free software; you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation; either version 2 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License along
#    with this program; if not, If not, see <http://www.gnu.org/licenses/>.


# Notes:
# As of the release of comtypes 1.1.2 the following fix is no longer applicable
#
# Once comtypes is installed, the following modifications
# need to be made for compatibility with ArcGIS 10.x:
# 1) Delete automation.pyc, automation.pyo, safearray.pyc, safearray.pyo
# 2) Add the following entry to the _ctype_to_vartype dictionary (near line 800):
#    POINTER(BSTR): VT_BYREF|VT_BSTR to automation.py

import os
import sys
from comtypes.client import CreateObject, GetModule, gen_dir
from comtypes import GUID

debug = 0  # change to 1 to output verbose traceback info

#todo add code to verify the install of comtypes supports the use of esriSystem.IArray

# Custom ERROR Handler Classes
class AOError(Exception):
    def __init__(self, msg='', severity=0, debug=0):
        self.severity = severity
        self.msg = msg
        self.debug = debug
        self.debugmsg = ''

    def getErr(self):
        #import sys
        import traceback
        if self.debug:
            tb = sys.exc_info()[2]
            tbinfo = traceback.format_tb(tb)[0]
            # Concatenate information together concerning the error into a message string
            self.debugmsg += "Traceback:\n" + tbinfo  # + "\n"

        self.debugmsg += "Error:\n" + self.msg
        return self.debugmsg


class ModuleError(AOError):
    pass


class LicenseError(AOError):
    pass


def _find_key(dic, val):
    """Return the the key for the first val found in a dictionary"""

    return [k for k, v in dic.iteritems() if v == val][0]


def ModuleImported(moduleName):
    """Return True if the module has already been imported"""

    try:
        if moduleName in sys.modules: 
            return True
        else:
            return False
    except:
        return False


def ModuleInstalled(moduleName):
    """Return True if the module can be imported"""

    try:
        exec 'import %s' % moduleName
        if ModuleImported(moduleName): 
            return True
        else:
            return False
    except:
        return False


def AddMsgAndPrint(msg, severity=0, gp=None, logger=None, ):
    """ Adds a GPmessage (in case this is run as a GP tool)
     and also prints the message to the screen (console standard output)
     and the passed python logger abject. 

    INPUT:
    msg (str): message to be output to the console and tool message window
                  logger (object,None): configured python logging object 
    severity {int,0}: severity of the output message

    OUTPUT:
    None

    NOTES:
    (1) The supported message severity are: 
            0 - For general informative messages (severity = 0)
            1 - For warning (severity = 1)
            2 - For errors (severity = 2)
    """

    print msg

    # Split the message on \n first, so that if it's multiple lines, 
    # a GPMessage and log message of the same severity will be added for
    # each line
    try:
        for msgstring in msg.split('\n'):
            if severity == 0:
                if gp: gp.AddMessage(msgstring)
                if logger: logger.info(msgstring)
            elif severity == 1:
                if gp: gp.AddWarning(msgstring)
                if logger: logger.warning(msgstring)
            elif severity == 2:
                if gp: gp.AddError(msgstring)
                if logger: logger.error(msgstring)
    except:
        pass


def MsgUser32(message="Hello world", title="ArcObjects Demo"):
    """Display a message using standard windows message dialog box
    """

    from ctypes import c_int, WINFUNCTYPE, windll
    from ctypes.wintypes import HWND, LPCSTR, UINT
    prototype = WINFUNCTYPE(c_int, HWND, LPCSTR, LPCSTR, UINT)
    fn = prototype(("MessageBoxA", windll.user32))
    return fn(0, message, title, 0)


# General COM Helper Functions
def GetTypeLibModule(typeLib, typeLibPath=None):
    """Generate a Python module (if not already present) from
    the typelibrary, containing interface classes, coclasses,
    constants, and structures without importing the module

    INPUT:
    typeLibFile (str, tuple): COM typelibrary file or GUID
    stypeLibPath (str): file path to the COM library file
 
    OUTPUT:
    Returns a wrapper for a COM typelibrary as a module object

    NOTES:
    (1) Defaults to searching the folder that contains with the
     calling python script
    (2) typeLib can be an ITypeLib COM pointer from a loaded type library, the pathname of a file
     containing a type library (.tlb, .exe or .dll), a tuple or list containing the GUID
     of a type library, a major and a minor version number, plus optionally a LCID, or any
     object that has a _reg_libid_ and _reg_version_ attributes specifying a type library.
    """

    #from comtypes.client import GetModule, gen_dir

    # tlib can be an ITypeLib COM pointer from a loaded typelibrary, the pathname of a file
    # containing a type library (.tlb, .exe or .dll), a tuple or list containing the GUID
    # of a typelibrary, a major and a minor version number, plus optionally a LCID, or any
    # object that has a _reg_libid_ and _reg_version_ attributes specifying a type library.
    try:
        if typeLibPath is None or typeLibPath == '':
            # looks like a GUID was passed or a *dll in the current folder
            pass
        else:
            # looks like name and path of a *dll or *olb file was passed
            typeLib = os.path.join(typeLibPath, typeLib)

        return GetModule(typeLib)

    except:
        raise ModuleError('The module could not be wrapped: %s\\%s' %(gen_dir, typeLib), 2, debug)


def NewObj(cClass, cInterface):
    """Creates a instance of COM object and returns a interface pointer.
    
    INPUT:
    comClass (COMType): specifies which object class to instantiate
    interface (COMType): specifies interface to the object class 
 
    OUTPUT:
    ptr (Pointer): interface pointer to the COM object

    NOTES:
    (1) Be aware that the COM Class and associated Interface
        are not always defined in the same COM Library
    """

    #from comtypes.client import CreateObject
    try:
        ptr = CreateObject(cClass, interface=cInterface)
        return ptr
    except:
        return None


def CType(cObj, cInterface):
    """Request query interface pointer from a COM obj reference.
    If the obj does not support the requested interface, None
    is returned.

    INPUT:
    cObj (COMType): specifies which COM object to query
    cInterface (COMType): specifies interface to cast the object 
 
    OUTPUT:
    ptr (Pointer): pointer to the casted COM query Interface(QI) object
    """

    try:
        return cObj.QueryInterface(cInterface)
    except:
        return None


def IArray():
    """Return a Array Object Interface reference"""

    # this is a small function that is a workaround to create
    # a interface reference to the ESRI Array COM Object
    esriTypeLibPath = GetESRITypeLibPath()
    esriSystem = GetTypeLibModule("esriSystem.olb", esriTypeLibPath)
    return NewObj("{8F2B6061-AB00-11D2-87F4-0000F8751720}", esriSystem.IArray)


def GetCLSID(cClass):
    """Return CLSID of MyClass as string"""

    return str(cClass._reg_clsid_)


# ESRI ArcObjects COM Helper Functions
def GetESRITypeLibPath():
    """Return folder location of ArcGIS COM type libraries (*.olb) files

    INPUT:
    None
    
    OUTPUT:
    comDir (str): directory containing the ESRI COM *.olb files

    NOTES:
    (1) looks in registry for the keypath of the esriFramework.olb 
        #HKEY_CLASSES_ROOT\TypeLib\{866AE5D3-530C-11D2-A2BD-0000F8774FB5}\a.0\0\win32
        #HKEY_CLASSES_ROOT\TypeLib\{866AE5D3-530C-11D2-A2BD-0000F8774FB5}\a.1\0\win32   
    """

    from _winreg import OpenKey, QueryValueEx, HKEY_CLASSES_ROOT

    keypath = r"TypeLib\{866AE5D3-530C-11D2-A2BD-0000F8774FB5}\a.1\0\win32"

    try:
        keyESRI = OpenKey(HKEY_CLASSES_ROOT, keypath)
        win32Dir = QueryValueEx(keyESRI, None)[0]
        comDir = os.path.dirname(win32Dir)

        if os.path.exists(comDir):
            return comDir
        else:
            raise AOError("Error determining location of ArcGIS COM library: %s" %(comDir), 2, debug)
    except:
        raise AOError("Unexpected error locating ArcGIS COM library!", 2, debug)


def WrapESRIDesktopTypeLib():
    """ Wrap the basic ArcGIS COM libraries for ArcGIS Desktop apps.
    This will not import the lib modules generated
    """

    esriTypeLibPath = GetESRITypeLibPath()
    GetTypeLibModule("esriFramework.olb", esriTypeLibPath)
    GetTypeLibModule("esriArcMapUI.olb", esriTypeLibPath)
    GetTypeLibModule("esriCatalogUI.olb", esriTypeLibPath)


def WrapESRIStandaloneTypeLib():
    """Wrap the commonly used ArcGIS COM libraries for standalone scripts.
    This will not import the lib modules generated
    """

    esriTypeLibPath = GetESRITypeLibPath()
    GetTypeLibModule("esriSystem.olb", esriTypeLibPath)
    GetTypeLibModule("esriGeometry.olb", esriTypeLibPath)
    GetTypeLibModule("esriCarto.olb", esriTypeLibPath)
    GetTypeLibModule("esriDisplay.olb", esriTypeLibPath)
    GetTypeLibModule("esriGeoDatabase.olb", esriTypeLibPath)
    GetTypeLibModule("esriDataSourcesGDB.olb", esriTypeLibPath)
    GetTypeLibModule("esriDataSourcesFile.olb", esriTypeLibPath)
    GetTypeLibModule("esriOutput.olb", esriTypeLibPath)


def WrapESRIAllTypeLib():
    """Wrap all *.olb type libraries found in the ESRI COM library path.
    This will not import the lib modules generated
    """

    esriTypeLibPath = GetESRITypeLibPath()
    moduleList = [os.path.basename(x) for x in os.listdir(esriTypeLibPath) if os.path.splitext(x)[1].lower() == '.olb']
    for module in iter(moduleList):
        GetTypeLibModule(module, esriTypeLibPath)


def InitESRIExtensionLicense(extension=None):		
    # Todo: need to add code to allow extension check-out
    pass


def InitESRIProductLicense(product='ArcView'):
    """Initialize ArcGIS license using ESRI COM objects, by default try to
    get 'ArcView'. If the passed product license name is not available try to get
    the next higher license.
    If a lower license is already initialized a error will be returned showing the
    initialized license name.
              
    INPUT:
    product (str,'ArcView'): specifies the license product code to
                             initialize
    OUTPUT:
    (str): specifies the license product code successfully initialized
           If requested or higher product code cannot be initialized
           a error is raised.

    NOTES:
    (1) The supported product codes are: 
         - 'ArcView'(default) : ArcGIS for Desktop Basic product code
         - 'Engine' : ArcGIS Engine Runtime product code
         - 'EngineGeoDB' : ArcGIS Engine Geodatabase Update product code
         - 'ArcEditor' : ArcGIS for Desktop Standard product code
         - 'ArcServer' : Server product code
         - 'ArcInfo' : ArcGIS for Desktop Advanced product code
    (2) This function is only used to initialize a license when executing a 
        stand-alone script and cannot change a existing initialized license
        in the current or separate process.
    (3) LicenseError will be raised if the requested or higher product code
        cannot be initialized successfully
    """
    from collections import OrderedDict
    #from comtypes import GUID

    esriTypeLibPath = GetESRITypeLibPath()

    esriSystem = GetTypeLibModule("esriSystem.olb", esriTypeLibPath)
    pInit = NewObj(esriSystem.AoInitialize, esriSystem.IAoInitialize)

    # libid:'6FCCEDE0-179D-4D12-B586-58C88D26CA78' is a raw_interface_only, no
    # implementation of VersionManager.dll
    guidTypeLib = (GUID("{6FCCEDE0-179D-4D12-B586-58C88D26CA78}"),1,0)
    esriVersionManager = GetTypeLibModule(guidTypeLib)
    pVM = NewObj(esriVersionManager.VersionManager, esriVersionManager.IArcGISVersion)
    vercode, ver, verpath = pVM.getActiveVersion()

    productCodeDict = {'ArcView': esriSystem.esriLicenseProductCodeBasic,
                       'ArcEditor': esriSystem.esriLicenseProductCodeStandard,
                       'ArcServer': esriSystem.esriLicenseProductCodeArcServer,
                       'ArcInfo': esriSystem.esriLicenseProductCodeAdvanced,
                       'Engine': esriSystem.esriLicenseProductCodeEngine,
                       'EngineGeoDB': esriSystem.esriLicenseProductCodeEngineGeoDB}

    # use the active version to update product codes for changes that occurred in 10.1
    if float(ver) < 10.1:
        # pre 10.1 license product codes
        productCodeDict['ArcView'] = esriSystem.esriLicenseProductCodeArcView
        productCodeDict['ArcEditor'] = esriSystem.esriLicenseProductCodeArcEditor
        productCodeDict['ArcInfo'] = esriSystem.esriLicenseProductCodeArcInfo

    # create a ordered dictionary of the product codes from
    # lowest to highest enumerated code values
    productDict = OrderedDict(sorted(productCodeDict.items(), key=lambda t: t[1]))

    # get the product code for the requested license
    if productDict.has_key(product):
        productCode = productDict[product]
    else:
        raise LicenseError("%s product code is not valid" %(product), 2, debug)

    # find the index of the product code in the ordered dict values
    indx = productDict.values().index(productCode)

    # verify a license is not already initialized
    # it's not possible to change to a different license once initialized
    try:
        initCode = pInit.InitializedProduct()
        if initCode >= productCode:
            # current initialized license is the same or higher than the requested license
            # good to go so return the product name
            return _find_key(productDict, initCode)
        else:
            # a lower license has already been initialized for the session and cannot be changed
            raise LicenseError('%s product license already initialized' %(_find_key(productDict, initCode)), 1, debug)
    except:
        # no product has been initialized, so ignore the error and boggie on
        pass

    # checkout the requested or higher available ESRI product license
    for code in productDict.values()[indx:]:
        licenseStatus = pInit.IsProductCodeAvailable(code)
        if licenseStatus != esriSystem.esriLicenseAvailable:
            # try the next higher available license
            continue
        licenseStatus = pInit.Initialize(code)
        if licenseStatus == esriSystem.esriLicenseCheckedOut:
            # return the product name
            return _find_key(productDict, code)
    raise LicenseError("Unable to initialize ArcGIS %s product license: %s" %(ver, product), 2, debug)


def GetESRIApp(app='ArcMap'):
    """Returns a application handle pointer. If the script is being run from
    inside the application boundary the parent application handle
    is returned, the passed param is ignored. If run outside the application boundary the first
    specified ArcGIS Desktop application instance is found and returned.
         
    INPUTS:
    app (str,'ArcMap'): specifies the running app to find the handle.

    OUTPUTS:
    pApp (Pointer): interface pointer to the application object,
    If a instance of the application is not found a error is raised.

    NOTES:
    (1) The supported apps are: 
         - 'ArcMap'
         - 'ArcCatalog'
    """
    appList = ["ArcMap", "ArcCatalog"]
    sESRITypeLibPath = GetESRITypeLibPath()
    esriFramework = GetTypeLibModule('esriFramework.olb', sESRITypeLibPath)
    esriArcMapUI = GetTypeLibModule('esriArcMapUI.olb', sESRITypeLibPath)
    esriCatalogUI = GetTypeLibModule('esriCatalogUI.olb', sESRITypeLibPath)

    # first check to see if we are running inside or outside a application boundary
    # and get the application handle
    pApp = NewObj(esriFramework.AppRef, esriFramework.IApplication)
    if pApp:
        # we are inside the application boundary, only ArcMap and ArcCatalog are supported
        if not (CType(pApp, esriCatalogUI.IGxApplication) or CType(pApp, esriArcMapUI.IMxApplication)):
            raise AOError('No \'%s\' application instances currently running.' %(app), 2, debug)
        else:
            return pApp
    else:
        # we are outside the application boundary so try to get a handle for the specified running app
        if app in appList:
            pAppROT = NewObj(esriFramework.AppROT, esriFramework.IAppROT)
            pApp = None
            iCount = pAppROT.Count
            for i in range(iCount):
                pApp = pAppROT.Item(i)
                if CType(pApp, esriCatalogUI.IGxApplication) and app == "ArcCatalog":
                    return pApp
                if CType(pApp, esriArcMapUI.IMxApplication) and app == "ArcMap":
                    return pApp
        else:
            raise AOError('No \'%s\' application instances currently running.' %(app), 2, debug)
