import camelot
import os,csv
import pandas as pd
import traceback
import sys
import pdfquery


pdfSourceDir = './pdf-sources'
# check if folder not exist
if not os.path.exists(pdfSourceDir):
    os.makedirs(pdfSourceDir)

def saveToCsv(results, fileName,parent):
    if not os.path.exists('./results'):
        os.makedirs('./results')
    # if parent contains / then create folder
    folderList = parent.split("/")
    parent = ""
    for folder in folderList:
        parent += folder+"/"
        if not os.path.exists('./results/'+parent):
            os.makedirs('./results/'+parent)
    
    if not os.path.exists('./results/'+parent+'/csv'):
        os.makedirs('./results/'+parent+'/csv')
    if not os.path.exists('./results/'+parent+'/excel'):
        os.makedirs('./results/'+parent+'/excel')
    csvFileName = './results/'+parent+'/csv/'+fileName+'.csv'
    with open(csvFileName, 'w', newline='') as csvfile:
        fieldnames = ['no', 'nama', 'jenis_kelamin', 'usia', 'rt', 'rw', 'nik', 'ket', 'nomor_tps', 'kelurahan_desa', 'kecamatan', 'kabupaten_kota', 'provinsi']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for result in results:
            writer.writerow(result)
    csvFile = pd.read_csv(csvFileName, encoding='cp1252')
    xlsxFileName = './results/'+parent+'/excel/'+fileName+'.xlsx'
    csvFile.to_excel(xlsxFileName, index=None, header=True)

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

def extractData(path,no,dpt):
    filename = path[path.rfind("/")+1:]
    results = []
    firstNo = no
    dataDoc = getDataDoc(path)
    tables=camelot.read_pdf(path,flavor='stream',pages='all')
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
            for index, row in df.iterrows():
                ok = False
                if row[1].strip() != "2" and row[1].strip() != "" and row[1].strip() != "NAMA" and row[1].strip() != "KABUPATEN/KOTA"  and row[1].strip() != 2 and row[2].strip() != "JENIS" and row[3].strip() != "JENIS" and row[3].strip() != "":
                    newDPT = dpt.copy()
                    newDPT["no"] = no
                    newDPT["ket"] = no
                    newDPT["nama"] = row[1]
                    if row[3].strip() == "L" or row[3].strip() == "P":
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
                            newDPT["ket"] = row[4]
                            splitRtRw = row[4].split("\n")
                            # get last and second last
                            if len(splitRtRw) > 2:
                                newDPT["rt"] = str(splitRtRw[len(splitRtRw)-1])
                                newDPT["rw"] = str(splitRtRw[len(splitRtRw)-2])

                    ok = True
                if ok:
                    results.append(newDPT)
                    no += 1
    except Exception as e:
        print(e,filename)
        print(traceback.format_exc())
        # or
        print(sys.exc_info()[2])
        # check if error folder not exist
        if not os.path.exists('./results/'+dpt["provinsi"]):
            os.makedirs('./results/'+dpt["provinsi"])
        if not os.path.exists('./results/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]):
            os.makedirs('./results/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"])
        if not os.path.exists('./results/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error"):
            os.makedirs('./results/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error")
        # copy file tp to error folder
        os.system("cp '"+path+"' './results/"+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error/"+filename+"'")

    if len(results) > 0:
        print(len(results),filename)
        saveToCsv(results, str(firstNo) +"_"+str((no-1))+ "_"+ filename,dpt["provinsi"]+"/"+dpt["kabupaten_kota"])
    return {
        "no":no,
        "results":results,
    }

def deepSearch(path,no,dpt):
    listFiles = os.listdir(path)
    for file in listFiles:
        if(os.path.isfile(path+"/"+file)):
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
        "rt":0,
        "rw":0,
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
        
        no = deepSearch("./pdf-sources/"+folderProvinsi+"/"+folderKabKota,no,dpt)
        print("Done "+kabupatenKota)
    print("Done "+folderProvinsi)
       
        
                


    