# pyArcObjects
Utility Functions for using fine grained ArcObjects in Python and GP script Tools.
This module only works on Windows platform. If ESRI ArcGIS installation is not found, an `AOError` exception is thrown.

Thanks to Mark Cederholm for providing inspiration and sharing ;-)

## Example usage:
```python
   try:
        #initializing a license only necessary if the standalone script
        pyArcObjects.InitESRIProductLicense('ArcView')
        esriGeoDatabase = pyArcObjects.GetTypeLib("esriGeoDatabase.olb",sESRITypeLibPath)
		esriDataSourcesGDB = pyArcObjects.GetTypeLib("esriDataSourcesGDB.olb",sESRITypeLibPath)
		# call some sample function code.......
        	sPath = "path/to/a/file_geodatbase.gdb"
		pWSF = pyArcObjects.NewObj(esriDataSourcesGDB.FileGDBWorkspaceFactory, \
					  esriGeoDatabase.IWorkspaceFactory)
		pWS = pWSF.OpenFromFile(sPath, 0)
		pDS = pyArcObjects.CType(pWS, esriGeoDatabase.IDataset)
		pyArcObjects.AddMsgAndPrint("Workspace name: " + pDS.BrowseName)
		pyArcObjects.AddMsgAndPrint("Workspace category: " + pDS.Category)
        
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

```

## Installation

You can pip install from GitHub:

    	pip install git+https://github.com/canopusgeo/pyArcObjects.git --trusted-host github.com