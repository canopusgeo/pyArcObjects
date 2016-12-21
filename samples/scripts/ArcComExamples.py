"""
Source Name:   ArcComExamples.py
Version:       ArcGIS 10.0 - 2012\02\16
Author:        Alex Catchpaugh 
Description:   Examples of functions using fine grained ArcObjects
               in Python. 
"""
#################### Standalone Examples ######################
def Standalone_ListAMDFiles():
    '''Example of using python and ArcObjects to access properties of a mosaicdataset.'''
    sESRITypeLibPath = pyArcObjects.GetESRITypeLibPath()
    esriSystem = pyArcObjects.GetTypeLibModule("esriSystem.olb",sESRITypeLibPath)
    esriGeoDatabase = pyArcObjects.GetTypeLibModule("esriGeoDatabase.olb",sESRITypeLibPath)
    esriDataSourcesGDB = pyArcObjects.GetTypeLibModule("esriDataSourcesGDB.olb",sESRITypeLibPath)
    esriDataSourcesRaster = pyArcObjects.GetTypeLibModule("esriDataSourcesRaster.olb",sESRITypeLibPath)

    sWSPath = r"I:\CGY\Geodata\raster_library\elevation\New_ECA_Elevation_Services.gdb"
    mDName = "ECA_WCB_LIDAR_BE_R"
    #Open the FGDB workspace
    pWSF = pyArcObjects.NewObj(esriDataSourcesGDB.FileGDBWorkspaceFactory, \
                  esriGeoDatabase.IWorkspaceFactory)
    pWS = pWSF.OpenFromFile(sWSPath, 0)

    #Use the helper class to find the MosaicWorkspaceExtension
    pMosaicWorkspaceExtensionHelper = pyArcObjects.NewObj(esriDataSourcesRaster.MosaicWorkspaceExtensionHelper, esriDataSourcesRaster.IMosaicWorkspaceExtensionHelper)
    pMosaicWorkspaceExtension = pMosaicWorkspaceExtensionHelper.FindExtension(pWS)

    #open an existing mosaic dataset
    pMosaicDataset = pMosaicWorkspaceExtension.OpenMosaicDataset(mDName)

    #get the feature class that represents the raster catalog component of the mosaic dataset
    pFeatureClass = pMosaicDataset.Catalog
 
    #create a new rastervalue object 
    pRasterValue = pyArcObjects.NewObj(esriGeoDatabase.RasterValue, \
                  esriGeoDatabase.IRasterValue2)
    
    #create a querfilter to return only the base rasters from the mosaicdataset 
    pQF = pyArcObjects.NewObj(esriGeoDatabase.QueryFilter, esriGeoDatabase.IQueryFilter)
    pQF.WhereClause = "CATEGORY = 1"
    pCursor = pFeatureClass.Search(pQF, True)
    pRow = pCursor.NextFeature()

    if not pRow:
        pyArcObjects.AddMsgAndPrint("Query returned no rows")
        return

    while pRow:
        valOID = pRow.Value(pCursor.FindField("OBJECTID"))
        valURIHash = pRow.Value(pCursor.FindField("URIHASH"))
        #this should never happen with the OBJECTID
        if valOID is None:
            pyArcObjects.AddMsgAndPrint("Null value")
            continue
        pFeature = pFeatureClass.GetFeature(valOID)
        if pFeature:
               pRasterCatalogItem = pyArcObjects.CType(pFeature, esriGeoDatabase.IRasterCatalogItem)
               #get the rasterdata set for the Catalog item
               pRDinfo = pyArcObjects.CType(pRasterCatalogItem.RasterDataset, esriGeoDatabase.IRasterDatasetInfo)
               #assign the rasterdataset to the rastervalue class to access some storage details 
               #pRasterValue.RasterDataset = pRasterCatalogItem.RasterDataset
               #print pRasterValue.RasterDatasetName.NameString
               print valURIHash
               print pRDinfo.FileList.Element(0)
        
        pRow = pCursor.NextFeature()

def Standalone_OpenFGDBWorkSpace():
    sESRITypeLibPath = pyArcObjects.GetESRITypeLibPath()
    esriGeoDatabase = pyArcObjects.GetTypeLib("esriGeoDatabase.olb",sESRITypeLibPath)
    esriDataSourcesGDB = pyArcObjects.GetTypeLib("esriDataSourcesGDB.olb",sESRITypeLibPath)
    sPath = "D:/Elevation_Image_Services.gdb"
    pWSF = pyArcObjects.NewObj(esriDataSourcesGDB.FileGDBWorkspaceFactory, \
                  esriGeoDatabase.IWorkspaceFactory)
    pWS = pWSF.OpenFromFile(sPath, 0)
    pDS = pyArcObjects.CType(pWS, esriGeoDatabase.IDataset)
    pyArcObjects.AddMsgAndPrint("Workspace name: " + pDS.BrowseName)
    pyArcObjects.AddMsgAndPrint("Workspace category: " + pDS.Category)
    return pWS

def Standalone_OpenSDEWorkspace():    
    sESRITypeLibPath = pyArcObjects.GetESRITypeLibPath()
    esriSystem = pyArcObjects.GetTypeLib("esriSystem.olb",sESRITypeLibPath)
    esriGeoDatabase = pyArcObjects.GetTypeLib("esriGeoDatabase.olb",sESRITypeLibPath)
    esriDataSourcesGDB = pyArcObjects.GetTypeLib("esriDataSourcesGDB.olb",sESRITypeLibPath)

    pPropSet = pyArcObjects.NewObj(esriSystem.PropertySet, esriSystem.IPropertySet)
    pPropSet.SetProperty("SERVER", "ss1392p")
    pPropSet.SetProperty("USER", "")
    pPropSet.SetProperty("INSTANCE", "5153")
    pPropSet.SetProperty("AUTHENTICATION_MODE", "DBMS")
    pPropSet.SetProperty("VERSION", "SDE.DEFAULT")
    pWSF = pyArcObjects.NewObj(esriDataSourcesGDB.SdeWorkspaceFactory, \
                  esriGeoDatabase.IWorkspaceFactory)
    pWS = pWSF.Open(pPropSet, 0)    
    pDS = pyArcObjects.CType(pWS, esriGeoDatabase.IDataset)
    pyArcObjects.AddMsgAndPrint("Workspace name: " + pDS.BrowseName)
    pyArcObjects.AddMsgAndPrint("Workspace category: " + pDS.Category)
    return pWS

def Standalone_OpenGISServer():
    sESRITypeLibPath = pyArcObjects.GetESRITypeLibPath()
    esriServer = pyArcObjects.GetTypeLib("esriServer.olb",sESRITypeLibPath)
    esriGeometry = pyArcObjects.GetTypeLib("esriGeometry.olb",sESRITypeLibPath)
 
    pServerConn = pyArcObjects.NewObj(esriServer.GISServerConnection, \
                         esriServer.IGISServerConnection)
    pServerConn.Connect("tuswpesri02")
    pServerManager = pServerConn.ServerObjectManager
    pServerContext = pServerManager.CreateServerContext("", "")

    #pUnk = pServerContext.CreateObject("esriGeometry.Polygon")
    pUnk = pServerContext.CreateObject(CLSID(esriGeometry.Polygon))
    pPtColl = pyArcObjects.CType(pUnk, esriGeometry.IPointCollection)
    XList = [0, 0, 10, 10]
    YList = [0, 10, 10, 0]
    iCount = 4
    for i in range(iCount):
        #pUnk = pServerContext.CreateObject("esriGeometry.Point")
        pUnk = pServerContext.CreateObject(CLSID(esriGeometry.Point))
        pPoint = pyArcObjects.CType(pUnk, esriGeometry.IPoint)
        pPoint.PutCoords(XList[i], YList[i])
        pPtColl.AddPoint(pPoint)
    pArea = pyArcObjects.CType(pPtColl, esriGeometry.IArea)
    print "Area = ", pArea.Area

    pServerContext.ReleaseContext()

def Standalone_QueryFGDBTableValues():
    sESRITypeLibPath = pyArcObjects.GetESRITypeLibPath()
    esriSystem = pyArcObjects.GetTypeLib("esriSystem.olb",sESRITypeLibPath)
    esriGeoDatabase = pyArcObjects.GetTypeLib("esriGeoDatabase.olb",sESRITypeLibPath)
    esriDataSourcesGDB = pyArcObjects.GetTypeLib("esriDataSourcesGDB.olb",sESRITypeLibPath)
 
    sPath = "c:/apps/Demo/Montgomery_full.gdb"
    sTabName = "Parcels"
    sWhereClause = "parcel_id = 6358"
    sFieldName = "zoning_s"
    
    pWSF = pyArcObjects.NewObj(esriDataSourcesGDB.FileGDBWorkspaceFactory, esriGeoDatabase.IWorkspaceFactory)
    pWS = pWSF.OpenFromFile(sPath, 0)
    pFWS = pyArcObjects.CType(pWS, esriGeoDatabase.IFeatureWorkspace)
    pTab = pFWS.OpenTable(sTabName)
    pQF = pyArcObjects.NewObj(esriGeoDatabase.QueryFilter, esriGeoDatabase.IQueryFilter)
    pQF.WhereClause = sWhereClause
    pCursor = pTab.Search(pQF, True)
    pRow = pCursor.NextRow()
    if not pRow:
        pyArcObjects.AddMsgAndPrint("Query returned no rows")
        return
    Val = pRow.Value(pTab.FindField(sFieldName))
    if Val is None:
        pyArcObjects.AddMsgAndPrint("Null value")
    
def Standalone_CreateFGDBTable():
    sESRITypeLibPath = pyArcObjects.GetESRITypeLibPath()
    esriSystem = pyArcObjects.GetTypeLib("esriSystem.olb",sESRITypeLibPath)
    esriGeoDatabase = pyArcObjects.GetTypeLib("esriGeoDatabase.olb",sESRITypeLibPath)
    esriDataSourcesGDB = pyArcObjects.GetTypeLib("esriDataSourcesGDB.olb",sESRITypeLibPath)

    sWSPath = "d:/temp/Lidar Backup.gdb"
    sTableName = "Test2224"
    
    pWSF = pyArcObjects.NewObj(esriDataSourcesGDB.FileGDBWorkspaceFactory, \
                  esriGeoDatabase.IWorkspaceFactory)
    pWS = pWSF.OpenFromFile(sWSPath, 0)
    pFWS = CType(pWS, esriGeoDatabase.IFeatureWorkspace)
    
    pOutFields = pyArcObjects.NewObj(esriGeoDatabase.Fields, esriGeoDatabase.IFields)
    pFieldsEdit = pyArcObjects.CType(pOutFields, esriGeoDatabase.IFieldsEdit)
    pFieldsEdit._FieldCount = 2
    pNewField = pyArcObjects.NewObj(esriGeoDatabase.Field, esriGeoDatabase.IField)
    pFieldEdit = pyArcObjects.CType(pNewField, esriGeoDatabase.IFieldEdit)
    pFieldEdit._Name = "OBJECTID"
    pFieldEdit._Type = esriGeoDatabase.esriFieldTypeOID
    pFieldsEdit._Field[0] = pNewField
    pNewField = pyArcObjects.NewObj(esriGeoDatabase.Field, esriGeoDatabase.IField)
    pFieldEdit = pyArcObjects.CType(pNewField, esriGeoDatabase.IFieldEdit)
    pFieldEdit._Name = "LUMBERJACK"
    pFieldEdit._Type = esriGeoDatabase.esriFieldTypeString
    pFieldEdit._Length = 50
    pFieldsEdit._Field[1] = pNewField
    pOutTable = pFWS.CreateTable(sTableName, pOutFields, \
                                 None, None, "")
    
    iField = pOutTable.FindField("LUMBERJACK")
    print "'LUMBERJACK' field index = ", iField
    pRow = pOutTable.CreateRow()
    pRow.Value[iField] = "I sleep all night and I work all day"
    pRow.Store()
    
####################  ArcMap Examples ######################
def ArcMap_GetSelectedGeometry():
    sESRITypeLibPath = pyArcObjects.GetESRITypeLibPath()
    esriFramework = pyArcObjects.GetTypeLib("esriFramework.olb",sESRITypeLibPath)
    esriArcMapUI = pyArcObjects.GetTypeLib("esriArcMapUI.olb",sESRITypeLibPath)
    esriSystem = pyArcObjects.GetTypeLib("esriSystem.olb",sESRITypeLibPath)
    esriCarto = pyArcObjects.GetTypeLib("esriCarto.olb",sESRITypeLibPath)
    esriGeoDatabase = pyArcObjects.GetTypeLib("esriGeoDatabase.olb",sESRITypeLibPath)
    esriGeometry = pyArcObjects.GetTypeLib("esriGeometry.olb",sESRITypeLibPath)

    # Get the application handle 
    pApp = pyArcObjects.GetESRIApp('ArcMap')
 
    # Get selected feature geometry
    pDoc = pApp.Document
    pMxDoc = CType(pDoc, esriArcMapUI.IMxDocument)
    pMap = pMxDoc.FocusMap
    pFeatSel = pMap.FeatureSelection
    pEnumFeat = CType(pFeatSel, esriGeoDatabase.IEnumFeature)
    pEnumFeat.Reset()
    pFeat = pEnumFeat.Next()
    if not pFeat:
        AddMsgAndPrint("No selection found.")
        return
    pShape = pFeat.ShapeCopy
    eType = pShape.GeometryType
    if eType == esriGeometry.esriGeometryPoint:
        pyArcObjects.AddMsgAndPrint("Geometry type = Point")
    elif eType == esriGeometry.esriGeometryPolyline:
        pyArcObjects.AddMsgAndPrint("Geometry type = Line")
    elif eType == esriGeometry.esriGeometryPolygon:
        pyArcObjects.AddMsgAndPrint("Geometry type = Poly")
    else:
        pyArcObjects.AddMsgAndPrint("Geometry type = Other")
    return pShape

def ArcMap_AddTextElement():
    sESRITypeLibPath = pyArcObjects.GetESRITypeLibPath()
    esriFramework = pyArcObjects.GetTypeLib("esriFramework.olb",sESRITypeLibPath)
    esriArcMapUI = pyArcObjects.GetTypeLib("esriArcMapUI.olb",sESRITypeLibPath)
    esriSystem = pyArcObjects.GetTypeLib("esriSystem.olb",sESRITypeLibPath)
    esriGeometry = pyArcObjects.GetTypeLib("esriGeometry.olb",sESRITypeLibPath)
    esriCarto = pyArcObjects.GetTypeLib("esriCarto.olb",sESRITypeLibPath)
    esriDisplay = pyArcObjects.GetTypeLib("esriDisplay.olb",sESRITypeLibPath)
    esriGeometry = pyArcObjects.GetTypeLib("esriGeometry.olb",sESRITypeLibPath)
    stdole = pyArcObjects.GetTypeLib("stdole.olb")

    # Get the application handle
    pApp = pyArcObjects.GetESRIApp('ArcMap')
    pFact = pyArcObjects.CType(pApp, esriFramework.IObjectFactory)

    # Get midpoint of focus map
    pDoc = pApp.Document
    pMxDoc = pyArcObjects.CType(pDoc, esriArcMapUI.IMxDocument)
    pMap = pMxDoc.FocusMap
    pAV = pyArcObjects.CType(pMap, esriCarto.IActiveView)
    pSD = pAV.ScreenDisplay
    pEnv = pAV.Extent
    dX = (pEnv.XMin + pEnv.XMax) / 2
    dY = (pEnv.YMin + pEnv.YMax) / 2
    pUnk = pFact.Create(CLSID(esriGeometry.Point))
    pPt = pyArcObjects.CType(pUnk, esriGeometry.IPoint)
    pPt.PutCoords(dX, dY)

    # Create text symbol
    pUnk = pFact.Create(CLSID(esriDisplay.RgbColor))
    pColor = pyArcObjects.CType(pUnk, esriDisplay.IRgbColor)
    pColor.Red = 255
    pUnk = pFact.Create(CLSID(stdole.StdFont))
    pFontDisp = pyArcObjects.CType(pUnk, stdole.IFontDisp)
    pFontDisp.Name = "Arial"
    pFontDisp.Bold = True
    pUnk = pFact.Create(CLSID(esriDisplay.TextSymbol))
    pTextSymbol = pyArcObjects.CType(pUnk, esriDisplay.ITextSymbol)
    pTextSymbol.Font = pFontDisp
    pTextSymbol.Color = pColor
    pTextSymbol.Size = 24
    pUnk = pFact.Create(CLSID(esriDisplay.BalloonCallout))
    pTextBackground = pyArcObjects.CType(pUnk, esriDisplay.ITextBackground)
    pFormattedTS = pyArcObjects.CType(pTextSymbol, esriDisplay.IFormattedTextSymbol)
    pFormattedTS.Background = pTextBackground

    # Create text element and add it to map
    pUnk = pFact.Create(CLSID(esriCarto.TextElement))
    pTextElement = pyArcObjects.CType(pUnk, esriCarto.ITextElement)
    pTextElement.Symbol = pTextSymbol
    pTextElement.Text = "Wink, wink, nudge, nudge,\nSay no more!"
    pElement = pyArcObjects.CType(pTextElement, esriCarto.IElement)
    pElement.Geometry = pPt
    
    pGC = pyArcObjects.CType(pMap, esriCarto.IGraphicsContainer)
    pGC.AddElement(pElement, 0)
    pGCSel = pyArcObjects.CType(pMap, esriCarto.IGraphicsContainerSelect)
    pGCSel.SelectElement(pElement)
    iOpt = esriCarto.esriViewGraphics + \
           esriCarto.esriViewGraphicSelection
    pAV.PartialRefresh(iOpt, None, None)

    # Get element width
    iCount = pGCSel.ElementSelectionCount
    pElement = pGCSel.SelectedElement(iCount - 1)
    pEnv = pyArcObjects.NewObj(esriGeometry.Envelope, esriGeometry.IEnvelope)
    pElement.QueryBounds(pSD, pEnv)
    pyArcObjects.AddMsgAndPrint("Width = ", pEnv.Width)
    return

def ArcMap_GetEditWorkspace():
    sESRITypeLibPath = pyArcObjects.GetESRITypeLibPath()
    esriSystem = pyArcObjects.GetTypeLib("esriSystem.olb",sESRITypeLibPath)
    esriEditor = pyArcObjects.GetTypeLib("esriEditor.olb",sESRITypeLibPath)
    esriGeoDatabase = pyArcObjects.GetTypeLib("esriGeoDatabase.olb",sESRITypeLibPath)
    pApp = pyArcObjects.GetESRIApp('ArcMap')
    pID = pyArcObjects.NewObj(esriSystem.UID, esriSystem.IUID)
    pID.Value = CLSID(esriEditor.Editor)
    pExt = pApp.FindExtensionByCLSID(pID)
    pEditor = pyArcObjects.CType(pExt, esriEditor.IEditor)
    if pEditor.EditState == esriEditor.esriStateEditing:
        pWS = pEditor.EditWorkspace
        pDS = pyArcObjects.CType(pWS, esriGeoDatabase.IDataset)
        pyArcObjects.AddMsgAndPrint("Workspace name: " + pDS.BrowseName)
        pyArcObjects.AddMsgAndPrint("Workspace category: " + pDS.Category)

def ArcMap_GetSelectedTable():
    sESRITypeLibPath = pyArcObjects.GetESRITypeLibPath()
    esriFramework = pyArcObjects.GetTypeLib("esriFramework.olb",sESRITypeLibPath)
    esriArcMapUI = pyArcObjects.GetTypeLib("esriArcMapUI.olb",sESRITypeLibPath)
    esriGeoDatabase = pyArcObjects.GetTypeLib("esriGeoDatabase.olb",sESRITypeLibPath)

    # Get the application handle
    pApp = pyArcObjects.GetESRIApp('ArcMap')
    pDoc = pApp.Document
    pMxDoc = pyArcObjects.CType(pDoc, esriArcMapUI.IMxDocument)
    pUnk = pMxDoc.SelectedItem
    if not pUnk:
        pyArcObjects.AddMsgAndPrint("Nothing selected.")
        return
    pTable = pyArcObjects.CType(pUnk, esriGeoDatabase.ITable)
    if not pTable:
        pyArcObjects.AddMsgAndPrint("No table selected.")
        return
    pDS = pyArcObjects.CType(pTable, esriGeoDatabase.IDataset)
    pyArcObjects.AddMsgAndPrint("Selected table: " + pDS.Name)

####################  ArcCatalog Examples ######################

def ArcCatalog_GetSelectedTable():
    sESRITypeLibPath = pyArcObjects.GetESRITypeLibPath()
    esriFramework = pyArcObjects.GetTypeLib("esriFramework.olb",sESRITypeLibPath)
    esriCatalogUI = pyArcObjects.GetTypeLib("esriCatalogUI.olb",sESRITypeLibPath)
    esriCatalog = pyArcObjects.GetTypeLib("esriCatalog.olb",sESRITypeLibPath)
    esriGeoDatabase = pyArcObjects.GetTypeLib("esriGeoDatabase.olb",sESRITypeLibPath)

    # Get the application handle
    pApp = pyArcObjects.GetESRIApp("ArcCatalog")
    pGxApp = pyArcObjects.CType(pApp, esriCatalogUI.IGxApplication)
    pGxObj = pGxApp.SelectedObject
    if not pGxObj:
        pyArcObjects.AddMsgAndPrint("Nothing selected.")
        return
    pGxDS = pyArcObjects.CType(pGxObj, esriCatalog.IGxDataset)
    if not pGxDS:
        pyArcObjects.AddMsgAndPrint("No dataset selected.")
        return
    eType = pGxDS.Type
    if not (eType == esriGeoDatabase.esriDTFeatureClass or eType == esriGeoDatabase.esriDTTable):
        pyArcObjects.AddMsgAndPrint("No table selected.")
        return
    pDS = pGxDS.Dataset
    pTable = pyArcObjects.CType(pDS, esriGeoDatabase.ITable)
    pyArcObjects.AddMsgAndPrint( "Selected table: " + pDS.Name)
 
############## Sample function calls with Error handeling ################    
import os
import sys
import traceback
import pyArcObjects

def main():

    try:
        #initializing a license only necessary if the standalone script
        pyArcObjects.InitESRIProductLicense('ArcView')
        pyArcObjects.InitESRIProductLicense('ArcView')
        #pyArcObjects.GetTypeLibModule('ertt', '')
        # call some sample function code.......
        #Standalone_ListAMDFiles()
        #Standalone_OpenFGDBWorkSpace()
        #Standalone_OpenSDEWorkspace()
        
    except pyArcObjects.AOError, err:
        # Output the python error messages to the script tool or Python Window
        # or a logging object
        pyArcObjects.AddMsgAndPrint(err.getErr(),2)
        
    except:
        #Unhandled exception occured, so here's the first traceback!"
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]

        # Concatenate information together concerning the error into a message string
        pymsg = "Traceback:\n" + tbinfo #+ "\n"
        pymsg += "Error:\n" + str(sys.exc_info()[1])

        # Return the python error messages to the script tool or Python Window
        pyArcObjects.AddMsgAndPrint(pymsg,2)

    finally:
        pass   

if __name__ == "__main__":
    main()