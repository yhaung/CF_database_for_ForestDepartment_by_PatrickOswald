import arcpy
from xlrd import open_workbook
import os
import img2pdf
from PIL import Image

txt_datasourceUTM = r"Data source: CF data from OMM based on official documents, alternative CF boundaries (where available, e.g. gps mapped or otherwise deemed more accurate) from RECOFT, 1:50.000 UTM topographic basemap from survey department"
txt_datasourceOI = r"Data source: CF data from OMM based on official documents, alternative CF boundaries (where available, e.g. gps mapped or otherwise deemed more accurate) from RECOFT, OneInch topographic basemap from survey department"

### get list of CF codes to process
import openpyxl
import re
xlsbook_cf2process = r"L:\OMM_projectMaster\mm_cfcertificates\products\cf_profiles\scripts\cfs2process.xlsx"
xlssheet_cf2process = 'cflist4process'
cfs = []
regex_cfcode = "^[0-9a-zA-Z]{9}_[a-zA-Z]+_\w+"

wb = openpyxl.load_workbook(xlsbook_cf2process)
sheet = wb.get_sheet_by_name(xlssheet_cf2process)
for cell in sheet['A']:
    try:
        if re.match(regex_cfcode,cell.value):
            cfs.append(cell.value)
            print(cell.value + " listed for CF profile creation.")
    except:
        print(str(cell.value) + ' does not look like a valid cf_code.')



pathdefaultpdfs = r"L:\OMM_projectMaster\mm_cfcertificates\products\cf_profiles\defaultpdfs"
pdf_missing_certmap = pathdefaultpdfs + "\\" + r"dummy_certmap_missing.pdf"
pdf_missing_satmap = pathdefaultpdfs + "\\" + r"dummy_satmap_missing.pdf"
pdf_missing_doccert = pathdefaultpdfs + "\\" + r"dummy_doccert_missing.pdf"
pdf_missing_docappl = pathdefaultpdfs + "\\" + r"dummy_docappl_missing.pdf"
pdf_missing_docvfv = pathdefaultpdfs + "\\" + r"dummy_docvfv_missing.pdf"
pdf_missing_docmngtpl = pathdefaultpdfs + "\\" + r"dummy_docmngtpl_missing.pdf"
pdf_missing_docfsr = pathdefaultpdfs + "\\" + r"dummy_docfsr_missing.pdf"

pdf_cover_certmap = pathdefaultpdfs + "\\" + r"cover_certmap.pdf"
pdf_cover_doccert = pathdefaultpdfs + "\\" + r"cover_doccert.pdf"
pdf_cover_docappl = pathdefaultpdfs + "\\" + r"cover_docappl.pdf"
pdf_cover_docvfv = pathdefaultpdfs + "\\" + r"cover_docvfv.pdf"
pdf_cover_docmngtpl = pathdefaultpdfs + "\\" + r"cover_docmngtpl.pdf"
pdf_cover_docfsr = pathdefaultpdfs + "\\" + r"cover_docfsr.pdf"
pdf_dummy_satmap_missing = pathdefaultpdfs + "\\" + r"dummy_satmap_missing.pdf"

xlsbook = r"L:\OMM_projectMaster\mm_cfcertificates\z_master_cfdb_omm\sourcedata\master_mm_cfcertificateFD_omm.xlsx"
xlssheet = r"cfCertificatesDB"
colxls_doccert = 'doc_cert'
colxls_certmap = 'certmap'
colxls_docmngtpl = 'doc_mngtpl'
colxls_docappl = 'doc_appl'
colxls_docfsr = 'doc_fsr'
colxls_docvfv = 'doc_vfv'
colxls_FDorgShp = 'FDorgShp'
colxls_l_doccert = 'l_doccert'
colxls_l_certmap = 'l_certmap'
colxls_l_docmngtpl = 'l_docmngt'
colxls_l_docappl = 'l_docappl'
colxls_l_docfsr = 'l_docfsr'
colxls_l_docvfv = 'l_docvfv'

cover_pdfMaps = ''
cover_pdfmngtpl = ''
cover_cert = ''
cover_appl = ''
dummy_pdfcertmap = ''
dummy_pdfmngtpl = ''



colxls_cCF = 'c_CF'



c_cf_focuscfs = list()

dfName_detail = 'detail'
dfName_overview = 'overview'
lyrName_googleSat = 'googleSat'
lyrName_OI = 'OI_topo'
lyrName_UTM = 'UTM_topo'
lytName_DDP = 'DDP'

lyrOvName_focusCF = 'focusCF'
lyrOvName_otherCF = 'otherCF'
lyrOvName_focusTownship = 'focusTownship'

col_c_cf = 'c_cf'



#==== prepare topo map mxd
filemxd = r"L:\OMM_projectMaster\mm_cfcertificates\products\cf_profiles\cfProfile.mxd"
pathexport = r"C:\tmp\CF\profiles20190920"
fileexport = "_CFprofile.pdf"
fileexport_mopt = "_mopt.pdf"
mxd = arcpy.mapping.MapDocument(filemxd)
df_detail = arcpy.mapping.ListDataFrames(mxd, dfName_detail)[0]
df_overview = arcpy.mapping.ListDataFrames(mxd, dfName_overview)[0]


#==== prepare cover page mxd
filecovermxd = r"L:\OMM_projectMaster\mm_cfcertificates\products\cf_profiles\cfProfileCoverPage.mxd"
dfName_coverframe = "coverframe"
lytName_coverDDP = 'DDP'

covermxd = arcpy.mapping.MapDocument(filecovermxd)
df_coverdetail = arcpy.mapping.ListDataFrames(covermxd, dfName_coverframe)[0]

mapExportsbyCF = {}

#============== functions
def xls2list(InputXlsBook, InputXlsSheet):
    dict_list = []
    book = open_workbook(InputXlsBook)
    sheet = book.sheet_by_name(InputXlsSheet)
    # reads 1st row for column names

    global colNames
    colNames = sheet.row_values(0)  # storing column names to be able to add them back in output excel file
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

#=============== start main script
#read cf master database
cfdb = xls2list(xlsbook, xlssheet)

arcpy.env.overwriteOutput = True


###start process for every CF in process list
for cf in cfs:
    query_cfs2export = '"c_cf" = \'' + cf + '\''
    c_cf_focuscf = cf
#update DDP layer with definition query
    lyr = arcpy.mapping.ListLayers(mxd, lytName_DDP, df_detail)[0]
    lyr.definitionQuery = query_cfs2export
# with arcpy.da.SearchCursor(lyr, col_c_cf) as cursor:
#     for row in cursor:
#         c_cf_focuscfs.append(row[0])
# del cursor

#========== start processing for each selected CF
# for c_cf_focuscf in c_cf_focuscfs:
    print('Processing started for CF: ' + c_cf_focuscf)
    query_focusCF = '"c_cf" IN (\'' + c_cf_focuscf + '\')'
    query_otherCF = '"c_cf" NOT IN (\'' + c_cf_focuscf + '\')'
    lyr = arcpy.mapping.ListLayers(mxd, lytName_DDP, df_detail)[0]
    lyr.definitionQuery = query_focusCF

    mxd.dataDrivenPages.refresh()

#---- configuring overview mapframe
    lyr = arcpy.mapping.ListLayers(mxd, lyrOvName_focusTownship, df_overview)[0]
    query_focusTS = '"TS_PCODE" = \'' + c_cf_focuscf[:9] + '\''
    lyr.definitionQuery = query_focusTS
    extent = lyr.getExtent(True)  # visible extent of layer
    df_overview.extend = extent
    arcpy.RefreshActiveView() # redraw the map

    sql = '"TS_PCODE" = \'' + c_cf_focuscf[:9] + '\''
    arcpy.SelectLayerByAttribute_management(in_layer_or_view=lyr, selection_type='NEW_SELECTION', where_clause=sql)
    # use the zoom to selected features method of the data frame to update the extent
    df_overview.zoomToSelectedFeatures()
    arcpy.SelectLayerByAttribute_management(lyr, "CLEAR_SELECTION")
    arcpy.RefreshActiveView()  # redraw the map

    # ---- setting up export file
    targetfile = pathexport + '\\' + c_cf_focuscf + fileexport
    finalPdf = arcpy.mapping.PDFDocumentCreate(targetfile)

    # ---- configuring cover page mapframe

    query_focusCF = '"c_cf" IN (\'' + c_cf_focuscf + '\')'
    lyr = arcpy.mapping.ListLayers(mxd, lytName_coverDDP, df_coverdetail)[0]
    lyr.definitionQuery = query_focusCF
    covermxd.dataDrivenPages.refresh()

    tmppdf = pathexport + '\\' + 'tmp_' + c_cf_focuscf + '_cover_' + fileexport
    arcpy.mapping.ExportToPDF(covermxd, tmppdf)
    finalPdf.appendPages(tmppdf)
    os.remove(tmppdf)

    # append certmap cover page
    if os.path.isfile(pdf_cover_certmap) == True:
        finalPdf.appendPages(pdf_cover_certmap)
    else:
        print('Certificate map coverpage not found.')


    # ---- configuring detail mapframe for UTM basemap
    lyr = arcpy.mapping.ListLayers(mxd, lyrName_OI, df_detail)[0]
    lyr.visible = False
    lyr = arcpy.mapping.ListLayers(mxd, lyrName_UTM, df_detail)[0]
    lyr.visible = True
    txtbox_datasource = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "txtbox_datasource")[0]
    txtbox_datasource.text = txt_datasourceUTM
    arcpy.RefreshActiveView()  # redraw the map

    tmppdf = pathexport + '\\' + 'tmp_' + c_cf_focuscf + '_utm_' + fileexport
    arcpy.mapping.ExportToPDF(mxd, tmppdf)
    finalPdf.appendPages(tmppdf)
    os.remove(tmppdf)

    # ---- configuring detail mapframe for OI basemap
    lyr = arcpy.mapping.ListLayers(mxd, lyrName_OI, df_detail)[0]
    lyr.visible = True
    lyr = arcpy.mapping.ListLayers(mxd, lyrName_UTM, df_detail)[0]
    lyr.visible = False
    txtbox_datasource = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "txtbox_datasource")[0]
    txtbox_datasource.text = txt_datasourceOI
    arcpy.RefreshActiveView()  # redraw the map

    tmppdf = pathexport + '\\' + 'tmp_' + c_cf_focuscf + '_oi_' + fileexport
    arcpy.mapping.ExportToPDF(mxd, tmppdf)
    finalPdf.appendPages(tmppdf)
    os.remove(tmppdf)


    for cf in cfdb:
        if cf[colxls_cCF] == c_cf_focuscf:  # filter to only process desired / current cf dataset
            # append satellite map to pdf
            if cf[colxls_certmap] in ['yes', 'incomplete'] or cf[colxls_FDorgShp] in ['yes', 'incomplete']:
                try:
                    # append sat image map pdf
                    tmppdf = pathexport + '\\' + c_cf_focuscf + '_sat.pdf'
                    if os.path.isfile(tmppdf) == True:
                        finalPdf.appendPages(tmppdf)
                    else:
                        finalPdf.appendPages(pdf_dummy_satmap_missing)
                        print('No satellite map for ' + c_cf_focuscf)
                except:
                    print('Problem importing satellite map image for ' + c_cf_focuscf)
            # append certificate map to pdf
            if cf[colxls_certmap] in ['yes', 'incomplete']:
                try:
                    # append sat image map pdf
                    tmppdf = pathexport + '\\' + 'tmp_' + c_cf_focuscf + '_certmap_' + fileexport
                    # opening image
                    image = Image.open(str(cf[colxls_l_certmap]))
                    # converting into chunks using img2pdf
                    pdf_bytes = img2pdf.convert(image.filename)
                    # opening or creating pdf file
                    file = open(tmppdf, "wb")
                    # writing pdf files with chunks
                    file.write(pdf_bytes)
                    # closing image file
                    image.close()
                    # closing pdf file
                    file.close()
                    finalPdf.appendPages(tmppdf)
                    os.remove(tmppdf)
                except:
                    print('Problem importing certificate map image for ' + c_cf_focuscf)
            else:
                print('No certificate map impage for ' + c_cf_focuscf)
                # append missing certmap cover page
                if os.path.isfile(pdf_missing_certmap) == True:
                    finalPdf.appendPages(pdf_missing_certmap)
                else:
                    print('Missing certificate map coverpage not found.')
            # append certificate pdf
            if cf[colxls_doccert] in ['yes','incomplete']:
                try:
                    # append certificate cover page
                    if os.path.isfile(pdf_cover_doccert) == True:
                        finalPdf.appendPages(pdf_cover_doccert)
                    else:
                        print('Certificate scan cover page not found.')
                    finalPdf.appendPages(cf[colxls_l_doccert])
                except:
                    print('Problem importing certificate pdf for ' + c_cf_focuscf)
            else:
                print('No certificate pdf for ' + c_cf_focuscf)
                # append missing certificate cover page
                if os.path.isfile(pdf_missing_doccert) == True:
                    finalPdf.appendPages(pdf_missing_doccert)
                else:
                    print('Missing certificate scan cover page not found.')
            # append management plan pdf
            if cf[colxls_docmngtpl] in ['yes', 'incomplete']:
                try:
                    # append management plan cover page
                    if os.path.isfile(pdf_cover_docmngtpl) == True:
                        finalPdf.appendPages(pdf_cover_docmngtpl)
                    else:
                        print('Management plan scan cover page not found.')
                    finalPdf.appendPages(cf[colxls_l_docmngtpl])
                except:
                    print('Problem importing management plan pdf for ' + c_cf_focuscf)
            else:
                print('No management plan pdf for ' + c_cf_focuscf)
                # append missing management plan scan cover page
                if os.path.isfile(pdf_missing_docmngtpl) == True:
                    finalPdf.appendPages(pdf_missing_docmngtpl)
                else:
                    print('Missing management plan scan cover page not found.')
            # append cf application pdf
            if cf[colxls_docappl] in ['yes', 'incomplete']:
                try:
                    # append CF application cover page
                    if os.path.isfile(pdf_cover_docappl) == True:
                        finalPdf.appendPages(pdf_cover_docappl)
                    else:
                        print('CF application scan cover page not found.')
                    finalPdf.appendPages(cf[colxls_l_docappl])
                except:
                    print('Problem importing CF application pdf for ' + c_cf_focuscf)
            else:
                print('No CF application pdf for ' + c_cf_focuscf)
                # append missing CF appliaction cover page
                if os.path.isfile(pdf_missing_docappl) == True:
                    finalPdf.appendPages(pdf_missing_docappl)
                else:
                    print('Missing CF application scan cover page not found.')
            # append field survey report pdf
            if cf[colxls_docfsr] in ['yes', 'incomplete']:
                try:
                    # append missing certificate cover page
                    if os.path.isfile(pdf_cover_docfsr) == True:
                        finalPdf.appendPages(pdf_cover_docfsr)
                    else:
                        print('Field survey report scan cover page not found.')
                    finalPdf.appendPages(cf[colxls_l_docfsr])
                except:
                    print('Problem importing field survey report pdf for ' + c_cf_focuscf)
            else:
                print('No field survey report pdf for ' + c_cf_focuscf)
                # append missing certificate cover page
                if os.path.isfile(pdf_missing_docfsr) == True:
                    finalPdf.appendPages(pdf_missing_docfsr)
                else:
                    print('Missing Field survey report scan cover page not found.')
            # append field survey report pdf
            if cf[colxls_docvfv] in ['yes', 'incomplete']:
                try:
                    # append Field survey Report scan cover page
                    if os.path.isfile(pdf_cover_docvfv) == True:
                        finalPdf.appendPages(pdf_cover_docvfv)
                    else:
                        print('VFV approval letter scan cover page not found.')
                    finalPdf.appendPages(cf[colxls_l_docvfv])
                except:
                    print('Problem importing VFV letter pdf for ' + c_cf_focuscf)
            else:
                print('No field VFV-letter pdf for ' + c_cf_focuscf)
                if cf[colxls_docvfv] not in ['na']:
                    # append missing vfv approval letter scan cover page
                    if os.path.isfile(pdf_missing_docvfv) == True:
                        finalPdf.appendPages(pdf_missing_docvfv)
                    else:
                        print('Missing VFV approval letter scan cover page not found.')
    finalPdf.saveAndClose()
    print("Profile as pdf created.")
    # try:
        # print (pdfsizeopt_cmd, targetfile, targetfile[:-4] + fileexport_mopt)
        # subprocess.call([pdfsizeopt_cmd, targetfile, targetfile[:-4] + fileexport_mopt])
    #
    # sub_cmds = "-sDEVICE=pdfwrite -dCompatibilityLevel=2 -dPDFSETTINGS=/ebook -dNOPAUSE -dQUIET -dBATCH -sOutputFile="
    # # input_pdf = input_dir + os.sep + input_pdf
    # # output_pdf = output_dir + os.sep + output_pdf
    # # ghostscript -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/printer -dNOPAUSE -dQUIET -dBATCH -sOutputFile=output.pdf input.pdf
    # input_pdf = targetfile
    # output_pdf = targetfile[:-4] + fileexport_mopt
    # # cmd_args += (gs_cmd, sub_cmds, output_pdf, input_pdf)
    # cmds = (gs_cmd + " " + sub_cmds + output_pdf + " " + input_pdf, shell=True)
    # subprocess.call(cmds)
    # print("Profile as reduced size pdf created.")
    # except:
    #     print("There was some error with the export of the reduced size pdf for the CF " + c_cf_focuscf)
    print('Processing completed for CF: ' + c_cf_focuscf)


arcpy.env.overwriteOutput = False

del mxd
del covermxd
print('Script finished running successfully.')