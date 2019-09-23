### Script for CF Mapping Layout Production
### created combined FD orig shapefiles and
### feature envelope for Data Driven Pages for May Layout

# switch to include checking and adding/populating CF code field FD orig source files in the data package
SWITCH_CHECKFDORIG = False   #true = FO orig shapefiles will be checked for c_cf field and populated if neccessary

import arcpy
import os
from xlrd import open_workbook

#CF excel master database and sheet name
xlsbook = r"L:\OMM_projectMaster\mm_cfcertificates\z_master_cfdb_omm\sourcedata\master_mm_cfcertificateFD_omm.xlsx"
xlssheet = r"cfCertificatesDB"
#full path to excel master sheet
xlsmastersheet_fullpath = xlsbook + '\\' + xlssheet + "$"
#essential column names in the CF excel master database
colxls_cCF = 'c_CF'
colxls_nm_suffix = 'nm_suffix'
colxls_FDorgShp = 'FDorgShp'



#column name in feature classes with the CF code
col_c_cf = 'c_cf'

#folder with the CF certificate folder packages
path_cf = r"L:\OMM_projectMaster\mm_cfcertificates\z_master_cfdb_omm\sourcedata\cf_permits"

#folder for temporary data created by and for this script
pathtmp = r"C:\tmp\CF\profiles\tmp"
#folder for output/resulting data/layers created by this script
pathoutputlayers = r"L:\OMM_projectMaster\mm_cfcertificates\products\cf_profiles\data\ProfileLayoutLayers"
#folder for script specific input data/dependencing
pathtemplates = r"L:\OMM_projectMaster\mm_cfcertificates\products\cf_profiles\data\ProfileLayoutLayers\templates"

#empty fc with the correct SRS and attributes where the FD original features are merged/copied to
file_cfBndFDCominedtemplate = pathtemplates + "\\" + "cfBndFDCominedtemplate.shp"
#feature class with the OMM certificate geometry for CF boundaries
#file_cfBndOMM = r"L:\OMM_projectMaster\mm_cfcertificates\z_master_cfdb_omm\product\mm_cfCertificatesFD_OMM_PY_20190909.shp"
file_cfBndOMM = r"L:\OMM_projectMaster\mm_cfcertificates\z_master_cfdb_omm\sourcedata\mm_cf_omm.gdb\master_mm_cf_omm\mm_cfCertificatesFD_omm_py"
#feature class with the OMM best available geometry for CF boundaries
file_cfBndBestAvailable = r"L:\OMM_projectMaster\mm_cfcertificates\z_master_cfdb_omm\sourcedata\mm_cf_omm.gdb\master_mm_cf_omm\mm_cfCertificatesBestAvailable_omm_py"


#temporary fc for combined FD orig. CF area polygons and OMM CF polygons
file_cfBndOMMFDCombined = pathtmp + "\\" + "cfBndOMMFDCombined.shp"
#temporary fc for combined FD orig. CF area polygons and OMM CF polygons dissolved by CF code
file_cfBndOMMFDCombinedDiss = pathtmp + "\\" + "cfBndOMMFDCombinedDiss.shp"
#temporary fc for polygons of the Envelopes of the combined FD orig. OMM CF polygons
file_cfBndOMMFDCombinedDissEnvelop = pathtmp + "\\" + "cfBndOMMFDCombinedDissEnvelop.shp"
#temporary fc for the combined FD orig features in one feature class
file_cfBndFDOrigCombined = pathtmp + "\\" + "mm_cfBndFDOrigCombined.shp"


#feature class of the combined and dissolved FD orig features to be used in Map Layouts
file_cfBndFDCombinedDiss = pathoutputlayers + "\\" + "mm_cfBndFDOrigCombinedDiss.shp"
#feature class of the feature envelope of the FD orig and OMM CF boundaries incl. all attributes from the excel CF database (used as Data Driven Page Layer in layout)
file_cfBndOMMFDComb4Envelop4DDP = pathoutputlayers + "\\" + "mm_cfBndOMMFDEnvelop4DDP.shp"
#feature class of the combined FD orig and OMM CF boundaries as points (used as Data Driven Page Layer in layout for inset map)
file_cfBndOMMFDCombinedDiss4DDP_PT = pathoutputlayers + "\\" + "mm_cfBndOMMFDCombinedDiss4DDP_PT"

#column name of the with of the CF boundary feature envelopes
COL_ENVELOPEWIDTH = 'MBG_Width'
#column name of the length of the CF boundary feature envelopes
COL_ENVELOPELENGH = 'MBG_Length'
#size of the mapframe in the layout (in meter)
MAPFRAME_SIZE = 19.5 / 100  #size in meter
#maximum targen scale in DDP for sat-maps (no feature shall be displayed at a map scale larger then this 1:5000 etc...)
MAXSCALE = 6000

#minimum target scale in DDP for topomaps background(no feature shall be displayed at a map scale smaller then this 1:30.000 etc...)
TOPOMINSCALE = 30000

#minimum target scale in DDP for covermap(no feature shall be displayed at a map scale smaller then this 1:30.000 etc...)
COVERMAPMINSCALE = 40000

#column name in the combined feature envelope layer used for attribute defined map scale for satmap background in ArcMaps Data Driven Pages
col_scaleOMMFD = 'scaleOMMFD'
#column name in the combined feature envelope layer used for attribute defined map scale for topomap background in ArcMaps Data Driven Pages
col_scaletopo = 'scaleTopo'
#column name in the combined feature envelope layer used for attribute defined map scale for Covermap background in ArcMaps Data Driven Pages
col_scalecover = 'scaleCover'

#size of layout dataframe in m (size in cm - some buffer around converted to m)
MAPFRAME_SIZE = (19.5 - 2) / 100   #make sure its float, not integer!
#size of layout dataframe of Covermap in m (size in cm - some buffer around converted to m)
COVERMAPFRAME_SIZE = (13.0 - 1) / 100   #make sure its float, not integer!

#Environment setting for Overwriting existing data
arcpy.env.overwriteOutput = True
#Environment setting for keeping original column names when doing a join
arcpy.env.qualifiedFieldNames = False

#============== functions
#read data from excel sheet into a dictionary list
def xls2list(InputXlsBook, InputXlsSheet):
    dict_list = []
    book = open_workbook(InputXlsBook)
    sheet = book.sheet_by_name(InputXlsSheet)

    keys = sheet.row_values(0)
    # read the rest rows for values and stores them in the list
    values = [sheet.row_values(i) for i in range(1, sheet.nrows)]
    for value in values:
        dict_list.append(dict(zip(keys, value)))
    if len(dict_list) > 0:
        return dict_list
    else:
        print ("Problem importing data from excel. Script will terminate.")
        sys.exit()

#reading CF excel master database
cfdb = xls2list(xlsbook, xlssheet)

#create a copy of the template fc in preparation to be populated with the polygons from the FD original shapefiles
arcpy.CopyFeatures_management(file_cfBndFDCominedtemplate, file_cfBndFDOrigCombined)




ic = arcpy.da.InsertCursor(file_cfBndFDOrigCombined,['SHAPE@', col_c_cf])
fcFDOrg2merge = list()
for cf in cfdb:
    print('Start working on ' + cf[colxls_cCF])
    #if the CF database has FD original shapefile in the package then
    #make sure this shapefile has an attribute c_cf and
    #that this attribute is filled with the cf code for the CF.
    #as not all features in that shapefile might be part of this CF (but maybe a different CF) this
    #script only populates the attribute c_cf automatically for those shapefiles where there is not even one feature with the
    #current CF code for this CF package
    if cf[colxls_FDorgShp] in ['yes']:
        fcFDorg = path_cf + '\\' + cf[colxls_cCF] + '_' + cf[colxls_nm_suffix] + '\\' + 'gis' + '\\' + cf[colxls_cCF] + '_FDorg.shp'
        if os.path.exists(fcFDorg):
            fcFDOrg2merge.append(fcFDorg)     #add to list for all FD orig shapefiles to eventually be merged into one combined fc
            #create CF code attribute field (and checking if it already exisits)
            if SWITCH_CHECKFDORIG == True:
                dsc = arcpy.Describe(fcFDorg)
                fields = dsc.fields
                fieldnames = [field.name for field in fields]  # get all field names except for the OID/FID field
                if col_c_cf not in fieldnames:
                    arcpy.AddField_management(fcFDorg, col_c_cf, "TEXT", field_length=50)
                    print('c_cf field added to ' + fcFDorg)

                #checking all features in that shapefile if there is at least one with the correct CF code for this package
                #if not --> all features will receive the CF code of the current CF in the c_cf field
                whereclause = '"' + col_c_cf + '" = \'' + cf[colxls_cCF] + '\''
                rows = [row for row in arcpy.da.SearchCursor(fcFDorg, col_c_cf, whereclause)]
                if len(rows) == 0:          #no polygon has a cf_code yet
                    with arcpy.da.UpdateCursor(fcFDorg,col_c_cf) as uc:
                        for row in uc:
                            row[0] = cf[colxls_cCF]
                            uc.updateRow(row)
                            print('Added ' + cf[colxls_cCF] + ' to ' + col_c_cf)
                    del uc
                else:
                    print('Some feature in ' + fcFDorg + ' already contains ' + cf[colxls_cCF] + ' in ' + col_c_cf + '\n No changes have been made tp that shapefile.')
            else:
                print('FD orig files checking disabled. FD orig shapefiles will be used as in the folders. This might result in an error if no c_cf field is yet created and poppulated for each FD orig shp.')
            # add features with the current CF code to the combined FD org shapefile
            rows2add = []
            whereclause = '"' + col_c_cf + '" = \'' + cf[colxls_cCF] + '\''
            with arcpy.da.SearchCursor(fcFDorg, ['SHAPE@', col_c_cf], whereclause) as sc:
                for row in sc:
                    rows2add.append((row[0], row[1]))
                    print (row[1] + ' will be added to FD')
            for row2add in rows2add:
                ic.insertRow(row2add)
            del sc
        else:
            print('File ' + fcFDorg + ' not found.')
    else:
        print("No FD orig shapefile for " + cf[colxls_cCF])
del ic  #delete insert querty
    #### end:  comment out if no FD orig files should be checked and modified

arcpy.Dissolve_management(file_cfBndFDOrigCombined, file_cfBndFDCombinedDiss, "c_cf", "", "MULTI_PART", "DISSOLVE_LINES")
print("Combined and dissolved CF FD orig shapefile exported to " + file_cfBndFDOrigCombined)
# merge OMM CF boundary file with combined FD orig data & CF best available bnd
arcpy.Merge_management([file_cfBndFDOrigCombined, file_cfBndOMM, file_cfBndBestAvailable], file_cfBndOMMFDCombined)
# dissolve all to get one feature per CF code
arcpy.Dissolve_management(file_cfBndOMMFDCombined, file_cfBndOMMFDCombinedDiss, "c_cf", "", "MULTI_PART", "DISSOLVE_LINES")
# save/convert combined polygons FO orig and OMM CF to point fc
arcpy.FeatureToPoint_management(file_cfBndOMMFDCombinedDiss, file_cfBndOMMFDCombinedDiss4DDP_PT, "INSIDE")
print("Combined FD orig and OMM CF areas as centerpoint (suitable for DDP) exported to " + file_cfBndOMMFDCombinedDiss4DDP_PT)
# create feature Envelope (box) for each CF and calculate dimensions
arcpy.MinimumBoundingGeometry_management(file_cfBndOMMFDCombinedDiss, file_cfBndOMMFDCombinedDissEnvelop, "ENVELOPE","LIST", "c_cf", "MBG_FIELDS")

#join attribute data from master excel list to feature envelope FC (this also includes attributes for CFs that are NOT yet in the OMM boundary FC
arcpy.MakeFeatureLayer_management(file_cfBndOMMFDCombinedDissEnvelop, 'lyr2join')
arcpy.AddJoin_management('lyr2join', col_c_cf, xlsmastersheet_fullpath, colxls_cCF, "KEEP_ALL")
#save joined feature envelope as FC that will be used as Data Driven Pages Layer
arcpy.CopyFeatures_management('lyr2join', file_cfBndOMMFDComb4Envelop4DDP)
print("Feature envelope with CF attributes (suitable for DDP) exported to " + file_cfBndOMMFDComb4Envelop4DDP)



#cleanup temporary files
#arcpy.arcpy.Delete_management(file_cfBndFDCombined)
#arcpy.arcpy.Delete_management(file_cfBndOMMFDCombinedDiss)
#arcpy.arcpy.Delete_management(file_cfBndOMMFDCombinedDissEnvelop)


#calculate scale field
#create attribute field for scale-value
arcpy.AddField_management(file_cfBndOMMFDComb4Envelop4DDP, col_scaleOMMFD, "LONG", 9)
arcpy.AddField_management(file_cfBndOMMFDComb4Envelop4DDP, col_scaletopo, "LONG", 9)
arcpy.AddField_management(file_cfBndOMMFDComb4Envelop4DDP, col_scalecover, "LONG", 9)
#determine maximum feature extend and create scale
with arcpy.da.UpdateCursor(file_cfBndOMMFDComb4Envelop4DDP, [COL_ENVELOPEWIDTH,COL_ENVELOPELENGH,col_scaleOMMFD, col_scaletopo, col_scalecover,col_c_cf]) as uc:
    for row in uc:
        maxsizefeature = row[0] if row[0] > row[1] else row[1]
        #set scale for sat map
        scale = 1 / (MAPFRAME_SIZE / maxsizefeature)   #convert from 1:2000 to 2000
        scale = MAXSCALE if scale < MAXSCALE else scale  #set maximum scale if feature is above threshold/too small
        scale = round(scale/100) * 100  #round to the nearest 100
        row[2] = scale

        # set scale for topo map
        scale = 1 / (MAPFRAME_SIZE / maxsizefeature)  # convert from 1:2000 to 2000
        scale = TOPOMINSCALE if scale < TOPOMINSCALE else scale  # set minimum scale if feature is above threshold/too large
        scale = round(scale / 100) * 100  # round to the nearest 100
        if scale > 35000:
            print("Mapscale for topomap layout >1:35.000 for CF" + row[5] + " Check for possible issues of multiple CFs in the FD-orig shapefile.")
        row[3] = scale
        # set scale for cover map
        scale = 1 / (COVERMAPFRAME_SIZE / maxsizefeature)  # convert from 1:2000 to 2000
        scale = COVERMAPMINSCALE if scale < COVERMAPMINSCALE else scale  # set minimum scale if feature is above threshold/too large
        scale = round(scale / 100) * 100  # round to the nearest 100
        row[4] = scale
        uc.updateRow(row)

print('Script finished running successfully.')