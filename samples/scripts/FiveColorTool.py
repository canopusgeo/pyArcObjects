# FiveColorTool.py

# Set up reporting

import traceback
import ArcComUtils
import sys

bTest = True # Set to True for standalone testing

sESRITypeLibPath  = ArcComUtils.GetESRITypeLibPath()
esriSystem = ArcComUtils.GetTypeLib("esriSystem.olb",sESRITypeLibPath)
esriGeometry = ArcComUtils.GetTypeLib("esriGeometry.olb",sESRITypeLibPath)
esriGeoDatabase = ArcComUtils.GetTypeLib("esriGeoDatabase.olb",sESRITypeLibPath)
esriGeoprocessing = ArcComUtils.GetTypeLib("esriGeoprocessing.olb",sESRITypeLibPath)

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
    
def Main():

    Status("Initializing...", -1)

    # Get parameters

    if bTest:
        sInClass = "c:/temp/New_Shapefile_Out2.shp"
        sColorFieldName = "COLOR_STR"
        sOverlapTolerance = ".01"
        sAssignRandom = ""
    else:
        sInClass = arcpy.GetParameterAsText(0)
        sColorFieldName = arcpy.GetParameterAsText(1)
        sOverlapTolerance = arcpy.GetParameterAsText(2)
        sAssignRandom = arcpy.GetParameterAsText(3) # default = False
    dTol = float(sOverlapTolerance)
    bRandomize = (sAssignRandom.lower() == "true")

    # Open feature class

    Status("Opening feature class...", -1)
    pGPUtil = ArcComUtils.NewObj(esriGeoprocessing.GPUtilities, \
                     esriGeoprocessing.IGPUtilities)
    pFC = pGPUtil.OpenFeatureClassFromString(sInClass)
    sShapeName = pFC.ShapeFieldName
    sOIDName = pFC.OIDFieldName
    sSubFields = sOIDName + "," + sShapeName
    iNumRecs = pFC.FeatureCount(None)

    # Set up value list

    Status("Checking color field...", -1)
    pTab = ArcComUtils.CType(pFC, esriGeoDatabase.ITable)
    iColorField = pTab.FindField(sColorFieldName)
    pField = pTab.Fields.Field(iColorField)
    pDomain = pField.Domain
    ValList = []
    iValCount = 5
    if pDomain:
        eType = pDomain.Type
        if eType == esriGeoDatabase.esriDTRange:
            pRangeDomain = ArcComUtils.CType(pDomain, esriGeoDatabase.IRangeDomain)
            MinVal = pRangeDomain.MinValue
            MaxVal = pRangeDomain.MaxValue
            DeltaVal = (MaxVal - MinVal) / (iValCount - 1)
            for i in range(iValCount):
                ValList.append(MinVal + (i * DeltaVal))
        elif eType == esriGeoDatabase.esriDTCodedValue:
            pCVDomain = ArcComUtils.CType(pDomain, esriGeoDatabase.ICodedValueDomain)
            iCount = pCVDomain.CodeCount
            if iCount < iValCount:
                sMsg = sColorFieldName + ": too few domain values"
                ArcComUtils.AddMsgAndPrint(sMsg,1)
                return
            if bRandomize:
                iValCount = iCount
            for i in range(iValCount):
                ValList.append(pCVDomain.Value(i))
        else:
            sMsg = sColorFieldName + ": unsupported domain type"
            ArcComUtils.AddMsgAndPrint(sMsg, 1)
            return
    elif pField.Type == esriGeoDatabase.esriFieldTypeString:
        ValList = ["A", "B", "C", "D", "E"]
    else:
        ValList = [0, 1, 2, 3, 4]

    # Build topology

    Status("Building neighbor table...", 0)
    NeighborTable = dict()
    pQF = ArcComUtils.NewObj(esriGeoDatabase.QueryFilter, esriGeoDatabase.IQueryFilter)
    pQF.SubFields = sSubFields
    pSF = ArcComUtils.NewObj(esriGeoDatabase.SpatialFilter, esriGeoDatabase.ISpatialFilter)
    pSF.GeometryField = sShapeName
    pSF.SubFields = sSubFields
    pSF.SpatialRel = esriGeoDatabase.esriSpatialRelIntersects
    #pSF.SpatialRel = esriGeoDatabase.esriSpatialRelOverlaps
    pFCursor = pFC.Search(pQF, True)
    iRecNum = 0
    while True:
        pFeat = pFCursor.NextFeature()
        if not pFeat:
            break
        iRecNum += 1
        if iRecNum % 100 == 0:
            sMsg = "Building neighbor table..." + str(iRecNum)
            iPos = int(iRecNum * 100 / iNumRecs)
            Status(sMsg, iPos)
        sOID = str(pFeat.OID)
        pShape = pFeat.Shape
        if not pShape:
            continue
        if pShape.IsEmpty:
            continue
        pPolygon = ArcComUtils.CType(pShape, esriGeometry.IPolygon)
        if pPolygon.ExteriorRingCount > 1:
            sMsg = "Multipart polygon detected at OID: " + sOID
            ArcComUtils.AddMsgAndPrint(sMsg, 1)
            return
        pTopoOp = ArcComUtils.CType(pShape, esriGeometry.ITopologicalOperator)
        pTopoOp.simplify()
        pSF.Geometry = pShape
        pSF.WhereClause = sOIDName + " <> " + sOID
        OIDList = []
        pFCursor2 = pFC.Search(pSF, True)
        while True:
            pFeat2 = pFCursor2.NextFeature()
            if not pFeat2:
                break
            pShape2 = pFeat2.Shape
            sOID2 = str(pFeat2.OID)
            # Skip point or multipoint intersection
            pIntShape = pTopoOp.Intersect(pShape2, esriGeometry.esriGeometry0Dimension)
            if not pIntShape.IsEmpty:
                continue
            # Check for polygon overlap
            pIntShape = pTopoOp.Intersect(pShape2, esriGeometry.esriGeometry2Dimension)
            if not pIntShape.IsEmpty:
                pArea = ArcComUtils.CType(pIntShape, esriGeometry.IArea)
                dArea = pArea.Area
                # If area within tolerance, treat as line intersection
                if dArea <= dTol:
                    OIDList.append(sOID2)
                    continue
                sMsg = "Polygon overlap detected at OID: " + sOID
                sMsg += "\n with OID: " + sOID2
                sMsg += "\nArea = " + str(dArea)
                ArcComUtils.AddMsgAndPrint(sMsg, 1)
                return
            OIDList.append(sOID2)
        del pFCursor2
        NeighborTable[sOID] = set(OIDList)
    del pFCursor

    # Analyze topology

    Status("Analyzing...", 0)
    NodeStack = []
    iLen = len(NeighborTable)
    iRecNum = 0
    while iLen > 0:
        
        iRecNum += 1
        if iRecNum % 1000 == 0:
            sMsg = "Analyzing..." + str(iRecNum)
            iPos = int(iRecNum * 100 / iNumRecs)
            Status(sMsg, iPos)

        # Search for Rule 1:
        # If X has fewer than 5 neighbors,
        # add X to the node stack and remove it from the topology

        bFound = False
        for sOID, OIDSet in NeighborTable.iteritems():
            if len(OIDSet) > 4:
                continue
            NodeStack.append([sOID, OIDSet])
            for sOID2 in OIDSet:
                OIDSet2 = NeighborTable[sOID2]
                OIDSet2.remove(sOID)
            del NeighborTable[sOID]
            bFound = True
            break
        if bFound:
            iLen = len(NeighborTable) # iLen -= 1
            continue

        # If no entries match rule 1, search for Rule 2:
        # If X has five neighbors, two of which have at most
        # at most seven neighbors and are not neighbors of each other,
        # add X to the node stack, remove it from the topology,
        # and combine the two neighbors into a single node connected
        # their neighbors plus X's remaining neighbors

        bFound = False
        for sOID, OIDSet in NeighborTable.iteritems():
            iCount = len(OIDSet)
            if iCount != 5:
                continue
            bFound = False
            OIDList = list(OIDSet)
            for i in range(iCount - 1):
                sOID2 = OIDList[i]
                OIDSet2 = NeighborTable[sOID2]
                if len(OIDSet2) > 7:
                    continue
                for j in range(i + 1, iCount):
                    sOID3 = OIDList[j]
                    if sOID3 in OIDSet2:
                        continue
                    OIDSet3 = NeighborTable[sOID3]
                    if len(OIDSet3) > 7:
                        continue
                    bFound = True
                    break
                if bFound:
                    break
            if not bFound:
                continue
            NodeStack.append([sOID, OIDSet])
            sN1 = sOID2
            sN2 = sOID3
            sNewOID = sN1 + "/" + sN2
            NewOIDSet = OIDSet | OIDSet2 | OIDSet3
            NewOIDSet.remove(sOID)
            NewOIDSet.remove(sN1)
            NewOIDSet.remove(sN2)
            del NeighborTable[sOID]
            del NeighborTable[sN1]
            del NeighborTable[sN2]
            for sOID2 in NewOIDSet:
                OIDSet2 = NeighborTable[sOID2]
                if sOID in OIDSet2:
                    OIDSet2.remove(sOID)
                if sN1 in OIDSet2:
                    OIDSet2.remove(sN1)
                if sN2 in OIDSet2:
                    OIDSet2.remove(sN2)
                OIDSet2.add(sNewOID)
            NeighborTable[sNewOID] = NewOIDSet
            break                
        if bFound:
            iLen = len(NeighborTable) # iLen -= 2
            continue
        # This should never happen
        # (if the overlap tolerance isn't too large)
        ArcComUtils.AddMsgAndPrint("Could not apply rule 2 at OID: " + sOID, 1)
        del NeighborTable
        del NodeStack
        return

    # Build color table

    Status("Building color table...", 0)
    ColorTable = dict()
    iLen = len(NodeStack)
    iRecNum = 0
    while iLen > 0:

        iRecNum += 1
        if iRecNum % 1000 == 0:
            sMsg = "Building color table..." + str(iRecNum)
            iPos = int(iRecNum * 100 / iNumRecs)
            Status(sMsg, iPos)

        NodeEntry = NodeStack.pop()
        sOID = NodeEntry[0]
        OIDSet = NodeEntry[1]
        ColorList = []
        for i in range(iValCount):
            ColorList.append(False)
        for sOID2 in OIDSet:
            sOID3 = sOID2.split("/")[0]
            if not sOID3 in ColorTable:
                continue
            i = ColorTable[sOID3]
            ColorList[i] = True
        Available = []
        for i in range(len(ColorList)):
            if ColorList[i]:
                continue
            Available.append(i)
        iCount = len(Available)
        if iCount == 0:
            # This should never happen
            # (unless the overlap tolerance is too large)
            sMsg = "Could not assign color at OID: " + sOID
            ArcComUtils.AddMsgAndPrint(sMsg, 1)
            del NodeStack
            del ColorTable
            return
        if bRandomize:
            i = random.randint(0, iCount - 1)
        else:
            i = 0
        iColor = Available[i]            
        for sOID2 in sOID.split("/"):
            ColorTable[sOID2] = iColor
        iLen -= 1

    # Calculate color field

    Status("Calculating color field...", 0)
    pQF.SubFields = sOIDName + "," + sColorFieldName
    pCursor = pTab.Update(pQF, False)
    iRecNum = 0
    while True:
        pRow = pCursor.NextRow()
        if not pRow:
            break
        iRecNum += 1
        if iRecNum % 1000 == 0:
            sMsg = "Calculating color field..." + str(iRecNum)
            iPos = int(iRecNum * 100 / iNumRecs)
            Status(sMsg, iPos)
        iOID = pRow.OID
        iColor = ColorTable[str(iOID)]
        pRow.Value[iColorField] = ValList[iColor]
        pCursor.UpdateRow(pRow)
    del pCursor
    del ColorTable
    del pTab
    del pFC
    Status("Done.", 0)

if __name__ == "__main__":
    try:
        ArcComUtils.InitESRIGPLicense('ArcEditor')
        import arcpy
        import random
        Main()

    except ArcComUtils.AOError, err:
        # Output the python error messages to the script tool or Python Window
        # or a logging object
        ArcComUtils.AddMsgAndPrint(err.getErr(),2)    

    except:
        #Unhandled exception occured, so here's the first traceback!"
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]

        # Concatenate information together concerning the error into a message string
        pymsg = "Traceback:\n" + tbinfo #+ "\n"
        pymsg += "Error:\n" + str(sys.exc_info()[1])

        # Return the python error messages to the script tool or Python Window
        ArcComUtils.AddMsgAndPrint(pymsg,2)
