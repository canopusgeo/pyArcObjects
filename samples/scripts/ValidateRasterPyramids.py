"""********************************************************************************************************************
TOOL NAME: UpdateRasterPathField
SOURCE NAME: UpdateRasterPathField.py
Tool Version: 1.0
ArcGIS VERSION: 10.0
AUTHOR: Alex Catchpaugh
REQUIRED ARGUMENTS: Input raster table, Output field name

TOOL DESCRIPTION: Export the rater columns sorce rasterdataset path to the
                  specified output field. The input raster table can be a Mosaic Dataset or Raster Catalog

Date Created: 01/28/2013
Updated: 
                 
Usage: UpdateRasterPathField Input_Raster_Table, Output FieldName, {where_clause}
*********************************************************************************************************************"""

#Importing standard library modules
import traceback
import ArcComUtils
import sys

bTest = True #False # Set to True for standalone testing

sESRITypeLibPath = ArcComUtils.GetESRITypeLibPath()
esriSystem = ArcComUtils.GetTypeLib("esriSystem.olb",sESRITypeLibPath)
esriGeoDatabase = ArcComUtils.GetTypeLib("esriGeoDatabase.olb",sESRITypeLibPath)
esriDataSourcesRaster = ArcComUtils.GetTypeLib("esriDataSourcesRaster.olb",sESRITypeLibPath)
esriGeoprocessing = ArcComUtils.GetTypeLib("esriGeoprocessing.olb",sESRITypeLibPath)

#==============================================================================
#   Methods
#==============================================================================

def Status(sMsg, pos=-1, gp=None):
    try:
        if gp is None:
            print sMsg
        elif pos == 0:
            gp.SetProgressor("step", sMsg)
        elif pos == -1:
            gp.SetProgressorLabel(sMsg)    
        else:        
            gp.SetProgressorLabel(sMsg)
            gp.SetProgressorPosition(pos)
    except Exception, e:
        #the user pressed the gp tools progress dialog "Cancel" so just re-raise exception
        if e.__class__.__name__ == 'ExecuteAbort' and str(e).lower() == 'cancelled function':
            raise 
        else:
            pass

def GetFieldNameIdx(pTable, fieldName):
    #get the position of the field that matches
    #the supplied name. If no raster fields
    #found or a error return None
    try:
        for i in range(pTable.Fields.FieldCount):
            if pTable.Fields.Field(i).Name == fieldName:
                fldIndex = i
                break
        return fldIndex
    except:
        return None

def GetRasterFieldName(pTable):
    #get the position of the first field with a
    #raster datatype in a table. If no raster fields
    #found or a error return None
    try:
        for i in range(pTable.Fields.FieldCount):
            if pTable.Fields.Field(i).Type == esriGeoDatabase.esriFieldTypeRaster:
                fldIndex = i
                break
        return fldIndex
    except:
        return None

def Main():

    #set up the default progressor
    Status("Initializing...", 0, arcpy.gp)

    #==============================================================================
    #   Assign the tool paramaters to variables
    #==============================================================================
    if bTest:
        #sInputMD = "D:/Mosaic_Datasets/Elevation_MosaicDatasets.gdb/ECA_WCB_LIDAR_BE_R"
        sInputMD =  r"D:\Mosaic_Datasets\Elevation_MosaicDatasets.gdb\ECA_WCB_LIDAR_FE_HS_R"
        #sInputMD = "D:/Elevation_DEV.gdb/test_managed"
        #sInputMD = "D:/OLD_Elevation_DEV.gdb/test_unmanaged"
        #sInputMD = "D:/Elevation_DEV.gdb/testmd_path"
        sWhereClause = "Category = 1" # exclude overview images
        sOutputFieldName = "FILENAME"

    else:
        sInputMD = arcpy.GetParameterAsText(0)
        sOutputFieldName = arcpy.GetParameterAsText(1)
        sWhereClause = arcpy.GetParameterAsText(2)
        arcpy.SetParameterAsText(3,sInputMD)
 
    #open the dataset
    pGPUtil = ArcComUtils.NewObj(esriGeoprocessing.GPUtilities, \
                     esriGeoprocessing.IGPUtilities)
    pDS = pGPUtil.OpenDatasetFromLocation(sInputMD)

    if not pDS:
        ArcComUtils.AddMsgAndPrint("Unable to open Dataset",1,arcpy.gp)
        return

    #Handle Mosaic Datasets and Raster Catalogs
    if pDS.Type == esriGeoDatabase.esriDTMosaicDataset: #'Mosaic Dataset'
        #cast the open dataset to the mosaic dataset interface
        pMosaicDataset = ArcComUtils.CType(pDS,esriDataSourcesRaster.IMosaicDataset) 
        #cast the dataset raster catalog component of the mosaic dataset type to a table interface
        pRasTable = ArcComUtils.CType(pMosaicDataset.Catalog, esriGeoDatabase.ITable)
        ArcComUtils.AddMsgAndPrint("Opened Mosaic Dataset...",0,arcpy.gp)

    elif pDS.Type == esriGeoDatabase.esriDTRasterCatalog: #'Raster Catalog'
        #cast the dataset raster catalog data type to a table interface
        pRasTable = ArcComUtils.CType(pDS, esriGeoDatabase.ITable)
        ArcComUtils.AddMsgAndPrint("Opened Raster Catalog...",0,arcpy.gp)

    elif pDS.Type == esriGeoDatabase.esriDTTable: #'Simple Table'
        #cast the dataset table data type to a table interface
        pRasTable = ArcComUtils.CType(pDS, esriGeoDatabase.ITable)
        ArcComUtils.AddMsgAndPrint("Opened Table...",0,arcpy.gp)

    else:
        ArcComUtils.AddMsgAndPrint("Dataset is not supported",2,arcpy.gp)
        return        

    #verify the the table has a raster column
    rasFieldIdx = GetRasterFieldName(pRasTable)
    if rasFieldIdx is None:
        ArcComUtils.AddMsgAndPrint("Raster column not found",2,arcpy.gp)
        return

    #verify the the table the output column name
    outputFieldIdx = GetFieldNameIdx(pRasTable,sOutputFieldName)
    if outputFieldIdx is None:
        ArcComUtils.AddMsgAndPrint("Output column name not found",2,arcpy.gp)
        return
    
    #create a queryfilter to return the rows that match the specified query catalog\MD
    #feature class 
    pTblQF = ArcComUtils.NewObj(esriGeoDatabase.QueryFilter, esriGeoDatabase.IQueryFilter)
    pTblQF.WhereClause = sWhereClause
    pTblCursor = pRasTable.Search(pTblQF, True)
    qryRowCnt = pRasTable.RowCount(pTblQF)
    pRow = pTblCursor.NextRow()

    if not pRow:
        ArcComUtils.AddMsgAndPrint("Query returned no rows" ,1 ,arcpy.gp)
        return

    pRowNum = 0.0

    #==============================================================================
    #   Begin processing raster table rows
    #==============================================================================    
    while pRow:
        pRowNum += 1.0
        Status("Processing rows...", int((pRowNum/qryRowCnt)*100) ,arcpy.gp)
        #Table class originated from a mosaic datasets catalog
        if pDS.Type == esriGeoDatabase.esriDTMosaicDataset:
            #There are two approaches to retrieve the file paths of the source rasters in the AMD, you can use the ItemPathsQuery
            #interface or the FunctionRasterDataset interface. They can produce differnt results and i really don't know why!

            ##setup the IItemPathQuery COM interface to return an array of file paths from the AMD URI column       
            #pItemQueryPath = ArcComUtils.CType(pMosaicDataset,esriDataSourcesRaster.IItemPathsQuery)

            #pQueryPathParam = ArcComUtils.NewObj(esriDataSourcesRaster.QueryPathsParameters, esriDataSourcesRaster.IQueryPathsParameters)
            #pItemQueryPath.QueryPathsParameters = pQueryPathParam
            #pQueryPathParam.BrokenPathsOnly = False
            #pQueryPathParam.FoldersOnly  = False
            #pQueryPathParam.PathDepth = 100
            #pQueryPathParam.QueryDatasetPaths = True
            #pQueryPathParam.QueryItemURIPaths = False
                 
            ##Get the current mosaic catalog row and return an array of paths as strings, using the ItemPathsQuery interface
            #pathArray = pItemQueryPath.GetItemPaths(pRow)

            #setup the IFunctionRasterDataset COM interface to return an array of file paths from the AMD Raster column
            #create a string array object to hold the raster paths
            pathArray = ArcComUtils.NewObj(esriSystem.StrArray, esriSystem.IStringArray)
            #get the functionRasterDataset for the current row
            pRasterValue = ArcComUtils.CType(pRow.Value(GetRasterFieldName(pRasTable)), esriGeoDatabase.IRasterValue2)
            pFunctionRasterDataset = ArcComUtils.CType(pRasterValue.RasterDataset, esriDataSourcesRaster.IFunctionRasterDataset)
            pMemberRasterDatasets = pFunctionRasterDataset.MemberRasterDatasets

            if pMemberRasterDatasets != None:                
                for i in range(pMemberRasterDatasets.Count):
                    pRasterDataset  = ArcComUtils.CType(pMemberRasterDatasets.Element(i), esriGeoDatabase.IRasterDataset2)
                    pathArray.Add(pRasterDataset.CompleteName)
            else:
                pathArray.Add('empty raster dataset')
 
        elif pDS.Type == esriGeoDatabase.esriDTRasterCatalog or pDS.Type == esriGeoDatabase.esriDTTable:
           #the input is a Raster Catalog or a Simple Table with a raster column
            try:
                #create a string array object to hold the raster paths
                pathArray = ArcComUtils.NewObj(esriSystem.StrArray, esriSystem.IStringArray)
                #get the raster column value for the current row
                pRasterValue = ArcComUtils.CType(pRow.Value(GetRasterFieldName(pRasTable)), esriGeoDatabase.IRasterValue2)
                if pRasterValue != None:
                    pRasterDataset = ArcComUtils.CType(pRasterValue.RasterDataset,esriGeoDatabase.IRasterDataset2)
                    pathArray.Add(pRasterDataset.CompleteName)
                else:
                    pathArray.Add('empty raster dataset')
            except:
                pathArray.Add('invalid raster dataset')

        #need to handle unsupported rasters made up of multiple files.
        if pathArray.Count > 1:
            ArcComUtils.AddMsgAndPrint('Multiple raster files per dataset are not supported, OID=%i' %(pRow.OID),1,arcpy.gp)
        elif pathArray.Count == 1:
            #update the row and commit the changes, trim the string to the field size
            pathString = pathArray.Element(0)
            #maxlength = pRow.Fields.Field(outputFieldIdx).Length
            #pRow.Value[outputFieldIdx] = pathString[:maxlength]
            #pRow.Store()

        if bTest:
            for j in range(pathArray.Count):
                pRasterPyramid = ArcComUtils.CType(pRasterDataset, esriDataSourcesRaster.IRasterPyramid3)
                ArcComUtils.AddMsgAndPrint('%s,%s,%i' %(pRow.OID,pathArray.Element(j),pRasterPyramid.PyramidLevel),0,arcpy.gp)
 
        pRow = pTblCursor.NextRow()
 
    ArcComUtils.AddMsgAndPrint("Completed processing...",0,arcpy.gp)        

    #==============================================================================
    #   Clean Up 
    #==============================================================================
    del pTblCursor        
    del pDS         

if __name__ == "__main__":
    try:
        #ArcComUtils.InitESRIGPLicense('ArcInfo')
        import arcpy
        Main()
        
    except ArcComUtils.AOError, err:
        # Output the python error messages to the script tool and\or Python Window
        # or a logging object
        ArcComUtils.AddMsgAndPrint(err.getErr(),2,arcpy.gp)    

    except:
        #Unhandled exception occured, so here's the first traceback!"
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]

        # Concatenate information together concerning the error into a message string
        pymsg = "Traceback:\n" + tbinfo #+ "\n"
        pymsg += "Error:\n" + str(sys.exc_info()[1])

        # Return the python error messages to the script tool and\or Python Window
        ArcComUtils.AddMsgAndPrint(pymsg,2,arcpy.gp)