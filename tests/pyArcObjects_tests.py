import unittest
import pyArcObjects

class pyArcObjectTestCase(unittest.TestCase):

    esriProductlist = ['Engine', 'EngineGeoDB', 'ArcServer', 'ArcView', 'ArcEditor',  'ArcInfo', ]

    def test_InitLicenseArcView(self):
        #test checkout of ArcView license
        product = 'ArcView'
        try:
            res = pyArcObjects.InitESRIProductLicense(product)
            # if the index of checked out license is greater than or equal to that requested than pass test
            indxres = self.esriProductlist.index(res)
            indxproduct = self.esriProductlist.index(product)
            self.assertTrue(indxres >= indxproduct)

        except pyArcObjects.LicenseError, err:
            self.assertTrue(False, msg=err.getErr())

        except Exception, err:
            self.assertTrue(False, err)


#todo: test the changes needed to numpy to support Array objects in AO

if __name__ == '__main__':
    unittest.main()
