# PublishPMF.py

# Set up reporting

import arcgisscripting
from ToolHelper import Status, Error, GetLibPath, NewObj, CType
from comtypes.client import GetModule
bTest = False # Set to True for standalone testing

sLibPath = GetLibPath();
GetModule(sLibPath + "esriSystem.olb")
GetModule(sLibPath + "esriCarto.olb")
GetModule(sLibPath + "esriPublisher.olb")

# Helper functions

def Init():
    import comtypes.gen.esriSystem as esriSystem
    pInit = NewObj(esriSystem.AoInitialize, esriSystem.IAoInitialize)
    eProduct = esriSystem.esriLicenseProductCodeArcView
##    eStatus = pInit.IsProductCodeAvailable(eProduct)
##    if eStatus != esriLicenseAvailable:
##        return False
##    eStatus = pInit.Initialize(eProduct)
##    if eStatus != esriLicenseCheckedOut:
##        return False
    # Check out Publisher
    eExtension = esriSystem.esriLicenseExtensionCodePublisher
    eStatus = pInit.IsExtensionCodeAvailable(eProduct, eExtension)
    if eStatus != esriSystem.esriLicenseAvailable:
        return False
    eStatus = pInit.CheckOutExtension(eExtension)
    if eStatus != esriSystem.esriLicenseCheckedOut:
        return False
    return True

def PublishPMF(pMapDoc, sOutPMF, bEnableLayout=True):
    import comtypes.gen.esriPublisher as esriPublisher
    # Configure publisher
    pPub = NewObj(esriPublisher.PublisherEngine, esriPublisher.IPMFPublish)
    pPropSet = pPub.GetDefaultPublisherSettings()
    pPropSet.SetProperty("DisplayFunctionalityMessage", False)
    pPropSet.SetProperty("OpenOnlyInArcReader", False)
    pPropSet.SetProperty("InternalObjectAccess", True)
    if bEnableLayout:
        pPropSet.SetProperty("DefaultViewType", esriPublisher.esriAPEViewTypeAll)
    else:
        pPropSet.SetProperty("DefaultViewType", esriPublisher.esriAPEViewTypeDataOnly)
    # Create PMF
    pLayout = pMapDoc.PageLayout
    pActiveView = pMapDoc.ActiveView
    bSuccess = True
    try:
        pPub.Publish(pLayout, pActiveView, pPropSet, False, sOutPMF)
    except:
        bSuccess = False
    return bSuccess

# Do it

def DoIt():
    import comtypes.gen.esriCarto as esriCarto
    gp = arcgisscripting.create(9.3)
    if bTest:
        g = None
    else:
        g = gp
    # Get parameters
    if bTest:
        sMxdPath = "c:/apps/demo/Hydrants.mxd"
        sPmfPath = "c:/apps/demo/Hydrants.pmf"
        sEnableLayout = ""
        gp.setproduct("ArcView")
    else:
        sMxdPath = gp.GetParameterAsText(0)
        sPmfPath = gp.GetParameterAsText(1)
        sEnableLayout = gp.GetParameterAsText(2) # default = True
    bEnableLayout = sEnableLayout.lower() != "false"
    Status("Checking out license...", -1, g)
    # Get Publisher license
    if not Init():
        Error("Could not get Publisher license.", g)
        return
    # Open mxd
    print "Opening MXD..."
    pMapDoc = NewObj(esriCarto.MapDocument, esriCarto.IMapDocument)
    try:
        pMapDoc.Open(sMxdPath)
    except:
        Error("Could not open MXD.", g)
        return
    # Publish PMF
    print "Publishing PMF..."
    bResult = PublishPMF(pMapDoc, sPmfPath, bEnableLayout)
    pMapDoc.Close()
    if not bResult:
        Error("Error publishing PMF.", g)
        return
    Status("Done.", -1, g)

if __name__ == "__main__":
    DoIt()
