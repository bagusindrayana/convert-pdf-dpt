import camelot
import os,csv
import pandas as pd
import traceback
import sys
import pdfquery
import shutil


pdfSourceDir = './results/fix-errors/pdf-sources'
resultsDir = './results/fix-errors/results'
deleteOriginal = False

# get arguments from --source --results --deleteOriginal
for i in range(1,len(sys.argv)):
    if sys.argv[i] == "--source":
        pdfSourceDir = sys.argv[i+1]
    if sys.argv[i] == "--results":
        resultsDir = sys.argv[i+1]
    if sys.argv[i] == "--deleteOriginal":
        deleteOriginal = str(sys.argv[i+1]).lower() == "true"



# check if folder not exist
if not os.path.exists(pdfSourceDir):
    os.makedirs(pdfSourceDir)

def saveToCsv(results, fileName,parent):
    if not os.path.exists(resultsDir):
        os.makedirs(resultsDir)
    # if parent contains / then create folder
    folderList = parent.split("/")
    parent = ""
    for folder in folderList:
        parent += folder+"/"
        if not os.path.exists(resultsDir+'/'+parent):
            os.makedirs(resultsDir+'/'+parent)
    
    if not os.path.exists(resultsDir+'/'+parent+'/csv'):
        os.makedirs(resultsDir+'/'+parent+'/csv')
    if not os.path.exists(resultsDir+'/'+parent+'/excel'):
        os.makedirs(resultsDir+'/'+parent+'/excel')
    csvFileName = resultsDir+'/'+parent+'/csv/'+fileName+'.csv'
    with open(csvFileName, 'w', newline='') as csvfile:
        fieldnames = ['no', 'nama', 'jenis_kelamin', 'usia', 'rt', 'rw', 'nik', 'ket', 'nomor_tps', 'kelurahan_desa', 'kecamatan', 'kabupaten_kota', 'provinsi']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for result in results:
            writer.writerow(result)
    csvFile = pd.read_csv(csvFileName, encoding='cp1252')
    xlsxFileName = resultsDir+'/'+parent+'/excel/'+fileName+'.xlsx'
    csvFile.to_excel(xlsxFileName, index=None, header=True)

def createTxtLog(path,fileName,log):
    folderList = path.split("/")
    parent = ""
    for folder in folderList:
        parent += folder+"/"
        if not os.path.exists('./'+parent):
            os.makedirs('./'+parent)
    txtFileName = parent+fileName+'.txt'
    with open(txtFileName, 'w', newline='') as txtfile:
        txtfile.write(log)

def getDataDoc(path):
    data = []
    tps = ""
    kelurahan = ""
    kecamatan = ""
    tpsIndex = 11
    pdf = pdfquery.PDFQuery(path)
    pdf.load()
    total = pdf.doc.catalog['Pages'].resolve()['Count']
    for i in range(0,total-1):
        try :
            checkTps = pdf.pq('LTPage[page_index="'+str(i)+'"] LTTextBoxHorizontal:contains("TPS")')
            
            if checkTps.text() == "TPS":
                tpsIndex = int(checkTps.attr("index"))
                result = pdf.pq('LTPage[page_index="'+str(i)+'"] LTTextBoxHorizontal[index="'+str(tpsIndex+1)+'"]')
                tps = result.text().replace(": ","")
            
            checkKelurahan = pdf.pq('LTPage[page_index="'+str(i)+'"] LTTextBoxHorizontal[index="'+str(tpsIndex-2)+'"]:contains("DESA/KELURAHAN")')
            if checkKelurahan.text() == "DESA/KELURAHAN":
                result = pdf.pq('LTPage[page_index="'+str(i)+'"] LTTextBoxHorizontal[index="'+str(tpsIndex-1)+'"]')
                kelurahan = result.text().replace(": ","")

            checkKecamatan = pdf.pq('LTPage[page_index="'+str(i)+'"] LTTextBoxHorizontal[index="'+str(tpsIndex-4)+'"]:contains("KECAMATAN")')
            if checkKecamatan.text() == "KECAMATAN":
                result = pdf.pq('LTPage[page_index="'+str(i)+'"] LTTextBoxHorizontal[index="'+str(tpsIndex-3)+'"]')
                kecamatan = result.text().replace(": ","")
            data.append({
                "tps":tps,
                "kelurahan":kelurahan,
                "kecamatan":kecamatan,
            })
        except Exception as e:
            print(e)
    return data

def checkDouble(dpt,results):
    for result in results:
        if result["nama"] == dpt["nama"] and result["jenis_kelamin"] == dpt["jenis_kelamin"] and result["usia"] == dpt["usia"] and result["rt"] == dpt["rt"] and result["rw"] == dpt["rw"] and result["nomor_tps"] == dpt["nomor_tps"] and result["kelurahan_desa"] == dpt["kelurahan_desa"]:
            return True
    return False

def extractData(path,no,dpt):
    filename = path[path.rfind("/")+1:]
    results = []
    firstNo = no
    dataDoc = getDataDoc(path)
    tables=camelot.read_pdf(path,flavor='stream',pages='all')
    haveError = False
    try:
        for table in tables:
            page = table.parsing_report['page']
            if(len(dataDoc) > page-1):
                dpt["nomor_tps"] = dataDoc[page-1]['tps']
                dpt["kelurahan_desa"] = dataDoc[page-1]['kelurahan']
                dpt["kecamatan"] = dataDoc[page-1]['kecamatan']
            if(dpt["nomor_tps"] == 0 or dpt["nomor_tps"] == "" or dpt["kelurahan_desa"] == "KELURAHAN" or dpt["kelurahan_desa"] == "" or dpt["kecamatan"] == "KECAMATAN" or dpt["kecamatan"] == ""):
                continue
            df = table.df.reset_index()  # make sure indexes pair with number of rows
            ok = False
            oldRT = ""
            oldRW = ""
            for index, row in df.iterrows():
                ok = False
                if row[1].strip() != "2" and row[1].strip() != "" and "NAMA" not in str(row[1]).strip() and "USIA" not in str(row[1]).strip() and row[1].strip() != "KABUPATEN/KOTA"  and row[1].strip() != 2 and row[2].strip() != "JENIS" and row[3].strip() != "JENIS" and row[3].strip() != "":
                    newDPT = dpt.copy()
                    newDPT["no"] = no
                    newDPT["nama"] = row[1].replace("/"," atau ").strip()
                    if row[2].strip() == "L" or row[2].strip() == "P":
                        newDPT["jenis_kelamin"] = row[2]
                    if len(str(row[3])) == 2:
                        newDPT["usia"] = row[3]
                    ok = True
                    try:
                        if row[1].strip().find("\nL") != -1 or row[1].strip().find("\nP") != -1:
                            splitRow1 = row[1].split("\n")
                            newDPT["nama"] = splitRow1[0].replace("/"," atau ").strip()
                            newDPT["jenis_kelamin"] = splitRow1[1]
                            newDPT["usia"] = row[2]
                            if len(row) > 3:
                                newDPT["ket"] = row[3]
                                splitRtRw = row[3].split("\n")
                                # get last and second last
                                if len(splitRtRw) > 2:
                                    newDPT["rt"] = str(splitRtRw[len(splitRtRw)-1])
                                    newDPT["rw"] = str(splitRtRw[len(splitRtRw)-2])
                                
                        elif row[3].strip() == "L" or row[3].strip() == "P":
                            newDPT["jenis_kelamin"] = row[3]
                            newDPT["usia"] = row[4]

                            try:
                                if len(str(row[5])) == 3:
                                    newDPT["rt"] = str(row[5])
                                    newDPT["rw"] = str(row[6])
                                elif len(str(row[6])) == 3 and len(str(row[7])) == 3:
                                    newDPT["rt"] = str(row[6])
                                    newDPT["rw"] = str(row[7])
                                elif len(str(row[8])) == 3:
                                    newDPT["rt"] = str(row[7])
                                    newDPT["rw"] = str(row[8])
                            except Exception:
                                if len(row) > 5:
                                    newDPT["ket"] = row[5]
                                    splitRtRw = row[5].split("\n")
                                    # get last and second last
                                    if len(splitRtRw) > 2:
                                        newDPT["rt"] = str(splitRtRw[len(splitRtRw)-1])
                                        newDPT["rw"] = str(splitRtRw[len(splitRtRw)-2])
                                    
                        else:
                            newDPT["jenis_kelamin"] = row[2]
                            newDPT["usia"] = row[3]
                            try:
                                if len(str(row[5])) == 3:
                                    newDPT["rt"] = str(row[5])
                                    newDPT["rw"] = str(row[6])
                                elif len(str(row[6])) == 3 and len(str(row[7])) == 3:
                                    newDPT["rt"] = str(row[6])
                                    newDPT["rw"] = str(row[7])
                                elif len(str(row[8])) == 3:
                                    newDPT["rt"] = str(row[7])
                                    newDPT["rw"] = str(row[8])
                            except Exception:
                                if len(row) > 4 and len(str(row[4]).strip()) > 8:
                                    newDPT["ket"] = row[4]
                                    splitRtRw = row[4].split("\n")
                                    # get last and second last
                                    if len(splitRtRw) > 2:
                                        newDPT["rt"] = str(splitRtRw[len(splitRtRw)-1])
                                        newDPT["rw"] = str(splitRtRw[len(splitRtRw)-2])
                                    elif len(splitRtRw) == 2:
                                        newDPT["rt"] = str(splitRtRw[0])
                                        newDPT["rw"] = str(splitRtRw[1])
                                    
   
                    except Exception as e:
                        print(filename,row)
                        print(filename,filename)
                        print(traceback.format_exc())
                        # or
                        print(sys.exc_info()[2])
                        haveError = True
                        createTxtLog(resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error",str(newDPT["nama"])+"_"+filename,"data DPT gagal di ekstrak : "+str(row[0])+","+str(row[1]))
                    if newDPT['rt'] == "" and newDPT['rw'] == "":
                        if len(row) > 4:
                            print(row[4],len(str(row[4])))
                            if len(str(row[4])) == 7:
                                newDPT["ket"] = row[3]
                                splitRtRw = row[4].split("\n")
                                if len(splitRtRw) > 1:
                                    newDPT["rt"] = str(splitRtRw[0])
                                    newDPT["rw"] = str(splitRtRw[1])
                            elif len(str(row[6])) > 8:
                                print(row[6])
                                if row[6][-3:].isdigit():
                                    # get the last 3 digit
                                    newDPT["rt"] = str(row[6][-3:])
                                newDPT["rw"] = str(row[7])
                            elif len(str(row[5])) > 8:
                                print(row[5])
                                splitRow = row[5].split(" ")
                                if len(splitRow) > 1 and str(splitRow[len(splitRow)-2]).strip().replace(".","")[0].isdigit() and str(splitRow[len(splitRow)-1]).strip().replace(".","")[0].isdigit():
                                    newDPT["rt"] = str(splitRow[0])
                                    newDPT["rw"] = str(splitRow[1])
                            elif len(str(row[4])) > 8:
                                print(row[4])
                                if row[4][-3:].isdigit():
                                    # get the last 3 digit
                                    newDPT["rt"] = str(row[4][-3:])
                                    newDPT["rw"] = str(row[6])
                        if newDPT['rt'] == "":
                            newDPT["rt"] = oldRT
                        if newDPT['rw'] == "":
                            newDPT["rw"] = oldRW

                        if newDPT['rt'] == "" and newDPT['rw'] == "":
                            print("Data RT RW not found ",row)
                            haveError = True
                            ok = False
                            createTxtLog(resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error",str(newDPT["nama"])+"_"+filename,"Data RT RW not found : "+str(row[0])+","+str(row[1]))
                        else:
                            oldRT = newDPT['rt']
                            oldRW = newDPT['rw']
                if ok:
                    if(checkDouble(newDPT,results)):
                        newDPT["ket"] =  str(newDPT["ket"])+str(no)
                    results.append(newDPT)
                    no += 1
    except Exception as e:
        print(e,filename)
        print(traceback.format_exc())
        # or
        print(sys.exc_info()[2])
        # check if error folder not exist
        if not os.path.exists(resultsDir+'/'+dpt["provinsi"]):
            os.makedirs(resultsDir+'/'+dpt["provinsi"])
        if not os.path.exists(resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]):
            os.makedirs(resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"])
        if not os.path.exists(resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error"):
            os.makedirs(resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error")
        # copy file tp to error folder
        # os.system("cp '"+path+"' './results/"+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error/"+filename+"'")
        shutil.copy(path, resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error/"+filename)

    if haveError:
        # os.system("cp '"+path+"' './results/"+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error/"+filename+"'")
        shutil.copy(path, resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error/"+filename)

    if len(results) > 0:
        print(len(results),filename)
        saveToCsv(results, str(firstNo) +"_"+str((no-1))+ "_"+ filename,dpt["provinsi"]+"/"+dpt["kabupaten_kota"])

        if deleteOriginal and haveError == False:
            print("Try Delete "+path)
            if os.path.exists(path):
                print("success Delete "+path)
                os.remove(path)
    return {
        "no":no,
        "results":results,
    }

def deepSearch(path,no,dpt):
    listFiles = os.listdir(path)
    for file in listFiles:
        if(os.path.isfile(path+"/"+file)):
            # if file pdf
            if file.endswith(".pdf"):
                print(file)
                extraxted = extractData(path+"/"+file,no,dpt)
                no = extraxted["no"]
        else:
            no = deepSearch(path+"/"+file,no,dpt)
    return no


# get folder list
folderList = os.listdir(pdfSourceDir)

no = 1
error_pdf = ""
for folderProvinsi in folderList:
    dpt = {
        "no":no,
        "nama":"Andi",
        "jenis_kelamin":"L",
        "usia":21,
        "rt":"",
        "rw":"",
        "nik":"-",
        "ket":"-",
        "nomor_tps":1,
        "kelurahan_desa":"KELURAHAN",
        "kecamatan":"KECAMATAN",
        "kabupaten_kota":"KABUPATEN",
        "provinsi":"PROVINSI",
    }
    folderKabupatenKotaList = os.listdir(pdfSourceDir+"/"+folderProvinsi)
    dpt["provinsi"] = folderProvinsi
    for folderKabKota in folderKabupatenKotaList:
        
        # remove "SALINAN DPT"
        kabupatenKota = folderKabKota.replace("SALINAN DPT","").replace("_"," ").strip()
        dpt["kabupaten_kota"] = kabupatenKota
        
        no = deepSearch(pdfSourceDir+'/'+folderProvinsi+"/"+folderKabKota,no,dpt)
        print("Done "+kabupatenKota)
    print("Done "+folderProvinsi)
       
        
                


    