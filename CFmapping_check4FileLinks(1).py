
#author:    Patrick Oswald (seastar3879@gmail.com)
#description:   checks if files and folders speficied in an excel table are valid/exist


from xlrd import open_workbook
from pathlib import Path
import re
import arcpy
# import xlwt


#http://www.blog.pythonlibrary.org/2014/03/24/creating-microsoft-excel-spreadsheets-with-python-and-xlwt/  excel and python
#http://www.blog.pythonlibrary.org/2014/03/24/creating-microsoft-excel-spreadsheets-with-python-and-xlwt/  excel and python


#variables
#name of excel work book with the input data
xlsBook =r"L:\OMM_projectMaster\mm_cfcertificates\forImport\20190716_CFfolderTemplate(New Version)\20190716_CFDBmasterTemplate.xlsx"
#name of excel sheet with the input data
xlsSheet =r"cfCertificatesDB"
fccfbnd = r"L:\OMM_projectMaster\mm_cfcertificates\forImport\20190716_CFfolderTemplate(New Version)\CF_Master.shp"
# fccfbnd = r"L:\OMM_projectMaster\mm_cfcertificates\forImport\mgw_CFpermitsFD_OMM\gis_data\cfCertificates_KZYL.shp"
col_cfcode_fc = "c_CF"  #field name of cf-code in feature class with cf area polygons

logfile = r"L:\OMM_projectMaster\mm_cfcertificates\forImport\20190716_CFfolderTemplate(New Version)\cfchecking.log"

#basefolder with CF datasets to check
basefolder = r"E:\YeHtetAung_CF_Database\mgw_CFpermitsFD_OMM\20190716_CFfolderTemplate(New Version)"

#regex to select specific records, column name + regular expression  in Excel file
#cffilter = ["D_SUBMIT","[\w]"]  #^MMR1002
cffilter = ["D_SUBMIT","YHA_2019-09-19"]  #^MMR1002NOK_2019-08-22

#Global variables
#column name with CF ID (used for reporting / finding record if there is an issue)
col_cCF = "c_CF"
col_nm_suffix = "nm_suffix"

col_doc_cert = "doc_cert"
col_doc_appl = "doc_appl"
col_doc_fsr = "doc_fsr"
col_doc_mngtpl = "doc_mngtpl"
col_doc_vfv = "doc_vfv"
col_certmap = "certmap"
col_FDorgShp = "FDorgShp"


colNames = []     #names of columns

issuelog = list()                   #- log for issues of missing docs that should exist
successlog = list()                 #- log of existing / found docs
incompletelog = list()              #- log for known missing or known incomplete docs
missingDocsExistlog = list()        #- log for false negative / false missing

def xls2list(InputXlsBook, InputXlsSheet):
    dict_list = []
    book = open_workbook(InputXlsBook)
    sheet = book.sheet_by_name(InputXlsSheet)
    #reads 1st row for column names

    global colNames
    colNames = sheet.row_values(0)  #storing column names to be able to add them back in output excel file
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

def isValidPath(id_datarow,path2check,topic):
    my_file = Path(path2check)
    if my_file.is_file() or my_file.is_dir():   #check for file
    # if my_file.is_dir(): # check for path
    #if my_file.exists(): check of file or path
    #if os.path.isfile(my_file) == True:  # pathname is a valid name
        successlog.append(id_datarow + ": " + topic + " found.")
        return True
    else:
        print (id_datarow, "File or folder ", path2check, "does not exist.")
        #logging.info(id_datarow, "File or folder ", path2check, "does not exist.")
        issuelog.append(id_datarow + ": " + topic + " missing.")
        return False


def negtestisValidFile(id_datarow,path2check,topic):
    my_file = Path(path2check)
    if my_file.is_file():   #check for file
    # if my_file.is_dir(): # check for path
    #if my_file.exists(): check of file or path
    #if os.path.isfile(my_file) == True:  # pathname is a valid name
        missingDocsExistlog.append(id_datarow + ": " + topic + " exist. Please check your excel table.")
        return True
    else:
        #print (id_datarow, "File or folder ", path2check, "does not exist.")
        #logging.info(id_datarow, "File or folder ", path2check, "does not exist.")
        incompletelog.append(id_datarow + ": " + topic + " still missing for a complete CF dataset.")
        return False



#
# def list2xls(outputrows):
#     book = xlwt.Workbook()
#     sheet = book.add_sheet(xlsSheet)
#     columns = list(outputrows[0].keys())  # list() is not need in Python 2.x
#     #style cells with date values as date
#     date_format = xlwt.XFStyle()
#     date_format.num_format_str = 'dd/MMM/yyyy'
#
#     for key, column in enumerate(colNames):
#         sheet.write(0, key, column)
#     for i, row in enumerate(outputrows):
#         for j, col in enumerate(columns):
#             if col in col_dateStyle:
#                 sheet.write(i+1, j, row[col],date_format)
#             else:
#                 sheet.write(i+1, j, row[col])
#     book.save(xlsBookOut)
#     print("Processing results exported to " + xlsBookOut)
#     logging.info("Processing results exported to " + xlsBookOut)



def exportlog(logfile, issues, incompletes, successes):
    file = open(logfile,"w")
    file.write("Logfile\n")

    file.write("Issues - Total:" + str(len(issues)) + "\n")
    issues.sort()
    for issue in issues:
        file.write(issue + "\n")
    # file.write("Issues: " + str(len(issues)-1) + "\n" )
    # print "Issues operations: " + str(len(issues)-1) + "\n"

    file.write("Indicated as missing but data seems to exist - Total:" + str(len(missingDocsExistlog)) + "\n")
    missingDocsExistlog.sort()
    for missingDocExistlog in missingDocsExistlog:
        file.write(missingDocExistlog + "\n")
    
    file.write("Incompete documents - Total:" + str(len(incompletes)) + "\n")
    incompletes.sort()
    for incomplete in incompletes:
        file.write(incomplete + "\n")

    file.write("Found docs - Total:"  + str(len(successes)) + "\n")
    successes.sort()
    for success in successes:
        file.write(success + "\n")
    file.close()
    print("Logfile exported to " + logfile)

def main():
    #import xls data to a list with dictionaries
    cfdb = list()
    cfdb = xls2list(xlsBook, xlsSheet)
    print(len(cfdb), "data rows imported.")

    #read polygones in cf featrue class
    cursor = arcpy.da.SearchCursor(fccfbnd, [col_cfcode_fc])
    fccf_codes = list()
    for row in cursor:
        fccf_codes.append(row[0])

    #start processing for each cf dataset
    cfcounter = 0
    for cf in cfdb:
        if len(cf[col_cCF]) > 0 and bool(re.search(cffilter[1], cf[cffilter[0]])):  #filter to only process desired cfdatasets
            cfcounter = cfcounter + 1  # to count total processed CF datasets
            codecfwithsuffix = cf[col_cCF] + "_" + cf[col_nm_suffix]    #cf code + camelized village name
            codecf = cf[col_cCF]  #cf-code
            #check for CF map
            setcodecf = {codecf}
            if cf[col_certmap].lower() == "yes":
                #test for jpg-scan
                path2check = basefolder + "\\" + codecfwithsuffix + "\\" + "gis" + "\\" + codecfwithsuffix + ".jpg"
                test = isValidPath(cf[col_cCF], path2check,"certificate map scan")
                #test for georef. file
                path2check = basefolder + "\\" + codecfwithsuffix + "\\" + "gis" + "\\" + codecfwithsuffix + "_wgs84.tif"
                test = isValidPath(cf[col_cCF], path2check, "georeferenced certificate map")
                #test for polygon in CF-feature class with matching c_cf
                # print (setcodecf)
                if setcodecf.issubset(set(fccf_codes)):
                    successlog.append(codecf + ": polygon with " + codecf +  " in feature class found.")
                else:
                    issuelog.append(codecf + ": no polygon with " + codecf +  " in feature class found.")
            else:
                path2check = basefolder + "\\" + codecfwithsuffix + "\\" + "gis" + "\\" + codecfwithsuffix + ".jpg"
                test = negtestisValidFile(cf[col_cCF], path2check, "certificate map scan")
                if setcodecf.issubset(set(fccf_codes)):
                    missingDocsExistlog.append(codecf + ": polygon with " + codecf +  " in feature class found while there is no (georeferenced) CF basemap.")
                else:
                    incompletelog.append(codecf + ": CF polygon not yet in CF GIS database.")

            # check for FD-orig shapefile
            path2check = basefolder + "\\" + codecfwithsuffix + "\\" + "gis" + "\\" + codecf + "_FDorg.shp"
            if cf[col_FDorgShp].lower() == "yes":
                test = isValidPath(cf[col_cCF], path2check, "FD original shapefile")
            else:
                test = negtestisValidFile(cf[col_cCF], path2check, "FD original shapefile")
                #incompletelog.append(codecf + ": original FD shapefile missing or incomplete.")

            # check for pdf-documents
            # test for certificate scan
            path2check = basefolder + "\\" + codecfwithsuffix + "\\" + "scan" + "\\" + codecfwithsuffix + "_cert.pdf"
            if cf[col_doc_cert].lower() == "yes":
                test = isValidPath(cf[col_cCF], path2check, "certificate")
            else:
                test = negtestisValidFile(cf[col_cCF], path2check, "certificate")

            # test for management plan scan
            path2check = basefolder + "\\" + codecfwithsuffix + "\\" + "scan" + "\\" + codecfwithsuffix + "_mngtpl.pdf"
            if cf[col_doc_mngtpl].lower() == "yes":
                test = isValidPath(cf[col_cCF], path2check, "management plan")
            else:
                test = negtestisValidFile(cf[col_cCF], path2check, "management plan")
                #incompletelog.append(codecf + ": management plan missing or incomplete.")

            # test for application scan
            path2check = basefolder + "\\" + codecfwithsuffix + "\\" + "scan" + "\\" + codecfwithsuffix + "_appl.pdf"
            if cf[col_doc_appl].lower() == "yes":
                test = isValidPath(cf[col_cCF], path2check, "application")
            else:
                test = negtestisValidFile(cf[col_cCF], path2check, "application")
                #incompletelog.append(codecf + ": application missing or incomplete.")

            #test for field survey report scan
            path2check = basefolder + "\\" + codecfwithsuffix + "\\" + "scan" + "\\" + codecfwithsuffix + "_fsr.pdf"
            if cf[col_doc_fsr].lower() == "yes":
                test = isValidPath(cf[col_cCF], path2check, "Field survey report")
            else:
                test = negtestisValidFile(cf[col_cCF], path2check, "Field survey report")
                #incompletelog.append(codecf + ": field survey report missing or incomplete.")

            # test for VFV approval scan
            path2check = basefolder + "\\" + codecfwithsuffix + "\\" + "scan" + "\\" + codecfwithsuffix + "_vfv.pdf"
            if cf[col_doc_vfv].lower() == "yes":
                test = isValidPath(cf[col_cCF], path2check, "VFV approval letter")
            elif cf[col_doc_vfv] == "na":  #as only CFs in RF/PPF need to have this, for the rest its not applicable
                test="dummy"  #do something if not applicable
            else:
                test = negtestisValidFile(cf[col_cCF], path2check, "VFV approval letter")
                #incompletelog.append(codecf + ": VFV land approval missing or incomplete.")

    exportlog(logfile, issuelog, incompletelog, successlog)   #export logfile
    print (cfcounter, "datasets processed.")
    print("Script finished running successfully.")


if __name__ == "__main__":
    main()


# to-do
# conmpare minimum bounding box with ordered points polygon
# if different --> flag
# check for geometries with 2 or less points --> flag









