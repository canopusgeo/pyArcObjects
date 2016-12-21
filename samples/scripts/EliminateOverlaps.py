# EliminateOverlaps.py
'''Takes a polygon feature class and copies it to a new feature class,
eliminating polygon overlaps. When two polygons overlap, the difference
is removed from the polygon with the larger area.

In Feature Class
The input polygon feature class.

Out Feature Class
The output feature class.

Maximum Sliver Area
The maximum allowable area for slivers (in coordinate system units).
Sliver polygons or polygon parts with this area or less will be removed.
'''

import traceback
import pyArcObjects
import sys

bTest = True # Set to True for standalone testing

sESRITypeLibPath  = pyArcObjects.GetESRITypeLibPath()
esriSystem = pyArcObjects.GetTypeLib("esriSystem.olb",sESRITypeLibPath)
esriGeometry = pyArcObjects.GetTypeLib("esriGeometry.olb",sESRITypeLibPath)
esriGeoDatabase = pyArcObjects.GetTypeLib("esriGeoDatabase.olb",sESRITypeLibPath)
esriGeoprocessing = pyArcObjects.GetTypeLib("esriGeoprocessing.olb",sESRITypeLibPath)

# Helper functions

def Status(sMsg, pos=-1):
    try:
        print sMsg
        if pos == -1:
            arcpy.setprogressorlabel(sMsg)
        elif pos == 0:
            arcpy.setprogressor("step", sMsg)
        else:        
            arcpy.setprogressorlabel(sMsg)
            arcpy.setprogressorposition(pos)
    except:
        pass
    
def RemoveSlivers(pShape, dTol):
    if pShape.IsEmpty:
        return False
    pGC = pyArcObjects.CType(pShape, esriGeometry.IGeometryCollection)
    iGeomCount = pGC.GeometryCount
    bSlivers = False
    for i in range (iGeomCount):
        j = iGeomCount - 1 - i
        pPart = pGC.Geometry(j)
        pPartArea = pyArcObjects.CType(pPart, esriGeometry.IArea)
        dArea = abs(pPartArea.Area)
        if dArea > dTol:
            continue
        pGC.RemoveGeometries(j, 1)
        bSlivers = True
    if not bSlivers:
        return False
    pTopoOp = pyArcObjects.CType(pShape, esriGeometry.ITopologicalOperator)
    pTopoOp.Simplify()
    return True

def Main():

    Status("Initializing...", -1)

    # Get parameters

    if bTest:
        sInClass = "c:/temp/New_Shapefile1.shp"
        sOutClass = "c:/temp/New_Shapefile_Out2.shp"
        sMaxSliverArea = "20"
    else:
        sInClass = arcpy.GetParameterAsText(0)
        sOutClass = arcpy.GetParameterAsText(1)
        sMaxSliverArea = arcpy.GetParameterAsText(2)
    dSliverTol = float(sMaxSliverArea)

    # Copy feature class

    Status("Copying feature class...", -1)
    arcpy.Copy_management(sInClass, sOutClass)

    # Repair geometry

    Status("Repairing geometry...", -1)
    arcpy.RepairGeometry_management(sOutClass)

    # Open feature class

    Status("Opening feature class...", -1)
    pGPUtil = pyArcObjects.NewObj(esriGeoprocessing.GPUtilities, esriGeoprocessing.IGPUtilities)
    pFC = pGPUtil.OpenFeatureClassFromString(sOutClass)
    sShapeName = pFC.ShapeFieldName
    sOIDName = pFC.OIDFieldName
    sSubFields = sOIDName + "," + sShapeName
    iNumRecs = pFC.FeatureCount(None)

    # Loop through records

    Status("Eliminating overlaps...", 0)
    pSF = pyArcObjects.NewObj(esriGeoDatabase.SpatialFilter, esriGeoDatabase.ISpatialFilter)
    pSF.GeometryField = sShapeName
    pSF.SubFields = sSubFields
    # (esriSpatialRelOverlaps fails if two polygons are identical)
    pSF.SpatialRel = esriGeoDatabase.esriSpatialRelIntersects
    # !!!! Do NOT assign subfields to an update cursor - script go boom !!!!
    pFCursor = pFC.Update(None, False)
    iRecNum = 0
    iUpdated = 0
    iDeleted = 0
    while True:

        # Get shape
        pFeat = pFCursor.NextFeature()
        if not pFeat:
            break
        iRecNum += 1
        if iRecNum % 100 == 0:
            sMsg = "Eliminating overlaps..." + str(iRecNum)
            iPos = int(iRecNum * 100 / iNumRecs)
            Status(sMsg, iPos)
        sOID = str(pFeat.OID)
        pShape = pFeat.ShapeCopy
        if not pShape:
            # This should never happen
            pFCursor.DeleteFeature()
            iDeleted += 1
            continue
        if pShape.IsEmpty:
            # This should never happen
            pFCursor.DeleteFeature()
            iDeleted += 1
            continue
        bUpdated = RemoveSlivers(pShape, dSliverTol)
        if pShape.IsEmpty:
            pFCursor.DeleteFeature()
            iDeleted += 1
            continue
        pTopoOp = pyArcObjects.CType(pShape, esriGeometry.ITopologicalOperator)
        pArea = pyArcObjects.CType(pShape, esriGeometry.IArea)
        dArea = pArea.Area

        # Get intersecting polygons
        
        bEmpty = False
        pSF.Geometry = pShape
        pSF.WhereClause = sOIDName + " <> " + sOID
        pFCursor2 = pFC.Search(pSF, True)
        pNewShape = None
        while True:
            pFeat2 = pFCursor2.NextFeature()
            if not pFeat2:
                break
            pShape2 = pFeat2.ShapeCopy
            sOID2 = str(pFeat2.OID)
            bResult = RemoveSlivers(pShape2, dSliverTol)
            if pShape2.IsEmpty:
                continue
            pIntShape = pTopoOp.Intersect(pShape2, esriGeometry.esriGeometry2Dimension)
            # Skip if no polygon overlap
            if pIntShape.IsEmpty:
                continue
            # Skip if other polygon is larger
            pArea2 = pyArcObjects.CType(pShape2, esriGeometry.IArea)
            dArea2 = pArea2.Area
            if dArea < dArea2:
                continue
            # Subtract adjacent polygon
            bUpdated = True
            pNewShape = pTopoOp.Difference(pShape2)
            bResult = RemoveSlivers(pNewShape, dSliverTol)
            bEmpty = pNewShape.IsEmpty
            if bEmpty:
                del pFeat2
                break
            pTopoOp = pyArcObjects.CType(pNewShape, esriGeometry.ITopologicalOperator)
            pArea = pyArcObjects.CType(pNewShape, esriGeometry.IArea)
            dArea = pArea.Area
        del pFCursor2
        if not bUpdated:
            continue
        if bEmpty:
            pFCursor.DeleteFeature()
            iDeleted += 1
            continue
        pFeat.Shape = pNewShape
        pFCursor.UpdateFeature(pFeat)
        iUpdated += 1

    del pFCursor
    del pFC

    sMsg = "Records updated: " + str(iUpdated)
    sMsg += "\nRecords deleted: " + str(iDeleted)
    pyArcObjects.AddMsgAndPrint(sMsg)

if __name__ == "__main__":
    try:
        pyArcObjects.InitESRIGPLicense('ArcEditor')
        import arcpy
        Main()

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

