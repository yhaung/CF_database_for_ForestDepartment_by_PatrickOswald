import openpyxl
import re

xlsbook_cf2process = r"L:\OMM_projectMaster\mm_cfcertificates\products\cf_profiles\scripts\cfs2process.xlsx"
xlssheet_cf2process = 'cflist4process'
cfs = []
regex_cfcode = "^[0-9a-zA-Z]{9}_[a-zA-Z]+_\w+"

wb = openpyxl.load_workbook(xlsbook_cf2process)
sheet = wb.get_sheet_by_name(xlssheet_cf2process)
for cell in sheet['A']:
    print(cell.value)
    try:
        if re.match(regex_cfcode,cell.value):
            cfs.append(cell.value)
    except:
        print(str(cell.value) + ' does not look like a valid cf_code.')

arcpy.env.overwriteOutput = True
#arcpy.env.qualifiedFieldNames = False

for c_cf_focuscf in cfs:
    dfName_detail = 'detail'
    dfName_overview = 'overview'
    lyrName_googleSat = 'googleSat'
    lytName_DDP = 'DDP'
    lyrOvName_focusCF = 'focusCF'
    lyrOvName_otherCF = 'otherCF'
    lyrOvName_township = 'focusTownship'
    filemxd = "CURRENT"
    pathexport = r"C:\tmp\CF\profiles20190920"
    fileexport = "_sat.pdf"

    # shpfdorigall = r"L:\OMM_projectMaster\mm_cfcertificates\z_master_cfdb_omm\product\mm_CFbndFDorigAllMerged_PY.shp"
    # shp_cf = r"L:\OMM_projectMaster\mm_cfcertificates\z_master_cfdb_omm\product\mm_cfCertificatesFD_OMM_PY_20190829.shp"
    # tmp_curFDorig = pathexport + '\\' + 'tmp_curFDorig.shp'
    # tmp_curOMM = pathexport + '\\' + 'tmp_curOMM.shp'

    fl_curFDorg = "fl_curFDorg"
    fl_curCFOMM = "fl_curCFOMM"
    # tmp_DDP = pathexport + '\\' + 'tmp_DDP.shp'
    # tmp_merged = pathexport + '\\' + 'tmp_mergedcurrent.shp'
    # tmp_joined = pathexport + '\\' + 'tmp_currentalljoined.shp'
    #
    #
    # filetmp_DDP = 'tmp_currentalljoined'

    mxd = arcpy.mapping.MapDocument(filemxd)
    df_detail = arcpy.mapping.ListDataFrames(mxd, dfName_detail)[0]
    df_overview = arcpy.mapping.ListDataFrames(mxd, dfName_overview)[0]
    query_focusCF = '"c_cf" IN (\'' + c_cf_focuscf + '\')'
    query_otherCF = '"c_cf" NOT IN (\'' + c_cf_focuscf + '\')'
    lyr = arcpy.mapping.ListLayers(mxd, lytName_DDP, df_detail)[0]

    # whereclause = '"c_cf" = \'' + c_cf_focuscf + '\''
    # rows = [row for row in arcpy.da.SearchCursor(shpfdorigall, ['c_cf'], whereclause)]
    # if len(rows) > 0:
    #     arcpy.MakeFeatureLayer_management(in_features=shpfdorigall, out_layer=fl_curFDorg, where_clause=whereclause)
    #     arcpy.MakeFeatureLayer_management(in_features=shp_cf, out_layer=fl_curCFOMM, where_clause=whereclause)
    #     arcpy.Merge_management([fl_curCFOMM, fl_curFDorg], tmp_merged)
    #     arcpy.Dissolve_management(tmp_merged, tmp_DDP, "c_cf", "", "MULTI_PART", "DISSOLVE_LINES")
    #     arcpy.JoinField_management(tmp_DDP, "c_cf", shp_cf, "c_cf")
    #     arcpy.CopyFeatures_management(tmp_DDP, tmp_joined)
    #     lyr.replaceDataSource(pathexport, "SHAPEFILE_WORKSPACE", filetmp_DDP)
    #
    lyr.definitionQuery = query_focusCF
    mxd.dataDrivenPages.refresh()

    pageID = mxd.dataDrivenPages.getPageIDFromName(c_cf_focuscf)
    mxd.dataDrivenPages.currentPageID = pageID

    lyr = arcpy.mapping.ListLayers(mxd, lyrOvName_township, df_overview)[0]
    query_focusTS = '"TS_PCODE" = \'' + c_cf_focuscf[:9] + '\''
    lyr.definitionQuery = query_focusTS
    lyr = arcpy.mapping.ListLayers(mxd, lyrOvName_township, df_overview)[0]
    ext = lyr.getExtent()
    df_overview.extent = ext
    mxd.dataDrivenPages.refresh()
    targetfile = pathexport + '\\' + c_cf_focuscf + fileexport
    arcpy.mapping.ExportToPDF(mxd, targetfile)
    # lyr = arcpy.mapping.ListLayers(mxd, lytName_DDP, df_detail)[0]
    # lyr.replaceDataSource(r"L:\OMM_projectMaster\mm_cfcertificates\z_master_cfdb_omm\product", "SHAPEFILE_WORKSPACE", "mm_cfCertificatesFD_OMM_PY_20190829")
    del mxd









#
#
#
#
# df_detail.scale = 30000 # we set the scale to 1:20,000
# mxd.activeView = df_detail.name
# arcpy.RefreshActiveView()
# arcpy.RefreshTOC()
#
#
#
# targetfile = pathexport + '\\' + c_cf_focuscf + fileexport
# arcpy.mapping.ExportToPDF(mxd, targetfile)
#
#

#
# import arcpy
# mxd = arcpy.mapping.MapDocument(r"C:\Project\ParcelAtlas.mxd")
# pageNameList = ["MPB", "PJB", "AFB", "ABB"]
# for pageName in pageNameList:
#     pageID = mxd.dataDrivenPages.getPageIDFromName(pageName)
#     mxd.dataDrivenPages.currentPageID = pageID
#     fieldValue = mxd.dataDrivenPages.pageRow.TSR  #example values from a field called TSR are "080102", "031400"
#     TRSTitle = arcpy.mapping.ListLayoutElements(MXD, "TEXT_ELEMENT", "TRSTitle")[0]
#     township, range, section = fieldValue[:2].strip("0"), fieldValue[2:-2].strip("0"), fieldValue[-2:].strip("0")
#     if section != "":
#         TRSTitle.text = "Section {0} T.{1}N. R.{2}W. W.M.".format(section, township, range)
#     else:
#         TRSTitle.text = "T.{0}N. R.{1}W. W.M.".format(township, range)
#     mxd.dataDrivenPages.printPages(r"\\olyfile\SUITE_303", "CURRENT")
# del mxd