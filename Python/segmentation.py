# -*- coding: utf-8 -*-
import arcpy

def Model():  # Model

    # To allow overwriting outputs change overwriteOutput option to True.
    arcpy.env.overwriteOutput = True

    CAMS = "CAMS"

if __name__ == '__main__':
    # Global Environment settings
    with arcpy.EnvManager(scratchWorkspace=r"C:\Users\samyan\Documents\ArcGIS\Projects\test_CAMS-segmentation\test_CAMS-segmentation.gdb", workspace=r"C:\Users\samyan\Documents\ArcGIS\Projects\test_CAMS-segmentation\test_CAMS-segmentation.gdb"):
        Model()
