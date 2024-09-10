import requests
import re
import random
import json
import sys
import os
import csv
import time
import pandas as pd
import datetime
from urllib.request import urlretrieve
import openpyxl
import shutil
import traceback
from fake_useragent import UserAgent
ua = UserAgent()

pdfSourceDir = './results/online'
resultsDir = './results/online-results'
deleteOriginal = False

for i in range(1,len(sys.argv)):
    if sys.argv[i] == "--source":
        pdfSourceDir = sys.argv[i+1]
    if sys.argv[i] == "--results":
        resultsDir = sys.argv[i+1]
    if sys.argv[i] == "--deleteOriginal":
        deleteOriginal = str(sys.argv[i+1]).lower() == "true"

def get_config():
    url = 'https://www.ilovepdf.com/pdf_to_excel'
    response = requests.get(url)
    html = response.text

    pattern = r'var ilovepdfConfig = (.*?);'
    pattern_task_id = r"ilovepdfConfig\.taskId = '(.*?)';"

    # Extract ilovepdfConfig as JSON string
    matches = re.search(pattern, html)
    json_string = matches.group(1) if matches else None
    data_config = json.loads(json_string) if json_string else {}

    # Extract taskId
    matches_task_id = re.search(pattern_task_id, html)
    data_task_id = matches_task_id.group(1) if matches_task_id else None

    # Select random server
    servers = data_config.get('servers', [])
    random_server = random.choice(servers) if servers else None

    return {
        'random_server': random_server,
        'data_task_id': data_task_id,
        'token': data_config.get('token')
    }

def download_excel(download_info):
    if not os.path.exists(download_info['download_path']):
        os.makedirs(download_info['download_path'])
    urlretrieve(download_info['url_download'], download_info['download_path']+"/"+download_info['file_name'])
    return download_info['download_path']+"/"+download_info['file_name']

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

def saveToCsv(results, fileName,parent):
    if not os.path.exists(resultsDir):
        os.makedirs(resultsDir)
    # if parent contains / then create folder
    folderList = parent.split("/")
    parent = ""
    for folder in folderList:
        parent += folder+"/"
        if not os.path.exists(resultsDir+"/"+parent):
            os.makedirs(resultsDir+"/"+parent)
    
    if not os.path.exists(resultsDir+"/"+parent+'/csv'):
        os.makedirs(resultsDir+"/"+parent+'/csv')
    if not os.path.exists(resultsDir+"/"+parent+'/excel'):
        os.makedirs(resultsDir+"/"+parent+'/excel')
    csvFileName = resultsDir+"/"+parent+'/csv/'+fileName+'.csv'
    with open(csvFileName, 'w', newline='') as csvfile:
        fieldnames = ['no', 'nama', 'jenis_kelamin', 'usia', 'rt', 'rw', 'nik', 'ket','alamat', 'nomor_tps', 'kelurahan_desa', 'kecamatan', 'kabupaten_kota', 'provinsi']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for result in results:
            writer.writerow(result)
    csvFile = pd.read_csv(csvFileName, encoding='cp1252')
    xlsxFileName = resultsDir+"/"+parent+'/excel/'+fileName+'.xlsx'
    csvFile.to_excel(xlsxFileName, index=None, header=True)

def convert_excel(file_path):
    config = get_config()
    data_task_id = config['data_task_id']
    random_server = config['random_server']
    token = config['token']

    # Extract the file name from the provided file path
    file_name = datetime.datetime.now().strftime("%Y%m%d%H%M%S") + "_" + os.path.basename(file_path)

    payload = {
        'name': file_name,
        'chunk': '0',
        'chunks': '1',
        'task': data_task_id,
        'preview': '1',
        'pdfinfo': '0',
        'pdfforms': '0',
        'pdfresetforms': '0',
        'v': 'web.0',
    }

    with open(file_path, 'rb') as f:
        files = [
            ('file', (file_name, f.read())),
        ]

    multipart_data = [(key, value) for key, value in payload.items()]

    headers = {
        'Accept-Language': 'en-GB,en-US;q=0.9,en;q=0.8',
        'Connection': 'keep-alive',
        'Origin': 'https://www.ilovepdf.com',
        'Referer': 'https://www.ilovepdf.com/',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-site',
        'User-Agent': ua.random,
        'accept': 'application/json',
        'authorization': f'Bearer {token}',
        'sec-ch-ua': '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }

    upload_url = f"https://{random_server}.ilovepdf.com/v1/upload"
    response = requests.post(upload_url, headers=headers, files=files, data=multipart_data)
    print("DOWNLOAD STATUS", response.status_code)
    json_response = response.json()

    payload_process = {
        'convert_to': 'xlsx',
        'output_filename': file_name + '.xlsx',
        'packaged_filename': 'ilovepdf_converted',
        'ocr': 0,
        'task': data_task_id,
        'tool': 'pdfoffice',
        'files[0][server_filename]': json_response['server_filename'],
        'files[0][filename]': file_name,
    }

    process_url = f"https://{random_server}.ilovepdf.com/v1/process"
    response_process = requests.post(process_url, headers=headers, data=payload_process)
    json_process = response_process.json()

    if json_process.get('status') == 'TaskSuccess':
        url_download = f"https://{random_server}.ilovepdf.com/v1/download/{data_task_id}"
        download_info = {
            "url_download": url_download,
            "file_name": file_name + ".xlsx",
            "download_path": "downloads"
        }
        download_excel(download_info)

        return {
            'success': True,
            'data': download_info,
        }
    else:
        return {
            'success': False,
            'data': json_process,
        }

def extractData(path,no,dpt):
    filename = path[path.rfind("/")+1:]
    results = []
    firstNo = no
    

    # Iterate the loop to read the cell values
    haveError = False
    results = []

    try:
        response = convert_excel(path)
        workbook = openpyxl.load_workbook(response['data']['download_path']+"/"+response['data']['file_name'])
        worksheets = workbook.worksheets
        
        for worksheet in worksheets:
            for row in worksheet.iter_rows(1, worksheet.max_row):
                data = {
                    "no":no,
                    "nama":"",
                    "jenis_kelamin":"",
                    "usia":"",
                    "rt":"",
                    "rw":"",
                    "nik":"",
                    "ket":"",
                    "alamat":"",
                    "nomor_tps":0,
                    "kelurahan_desa":"KELURAHAN",
                    "kecamatan":"KECAMATAN",
                    "kabupaten_kota":"KABUPATEN",
                    "provinsi":"PROVINSI",
                }
                cels = []
                for cell in row:
                    cels.append(cell.value)
                if len(cels) > 7:
                    if cels[0] != None and str(cels[0]) != "NO" and str(cels[0]) != "1" and str(cels[1]) != "NAMA" and str(cels[1]) != "2" and "PROVINSI" not in str(cels[0]) and "DAFTAR PEMILIH" not in str(cels[0]) and "Rekapitulasi" not in str(cels[0]): 
                        data['no'] = no
                        data['nama'] = cels[1]
                        data['jenis_kelamin'] = cels[2]
                        data['usia'] = cels[3]
                        data['alamat'] = cels[4]
                        data['rt'] = cels[5]
                        data['rw'] = cels[6]
                        data['ket'] = cels[7]
                    elif cels[len(row)-2] != None and cels[len(row)-2] != None and "KECAMATAN" in str(cels[len(row)-2]) and "KELURAHAN TPS" in str(cels[len(row)-2]):
                        valueSplit = str(cels[len(row)-1]).strip().split(":")
                        dpt["nomor_tps"] = valueSplit[3]
                        dpt['kelurahan_desa'] = valueSplit[2].strip()
                        dpt['kecamatan'] = valueSplit[1].strip()

                    if data['nama'] != None and data['nama'] != "":
                        newDPT = dpt.copy()
                        newDPT['no'] = no
                        newDPT['nomor_tps'] =  data["nomor_tps"]
                        newDPT['nama'] = data["nama"]
                        newDPT['jenis_kelamin'] = data["jenis_kelamin"]
                        newDPT['usia'] = data["usia"]
                        newDPT['alamat'] = str(data["alamat"]).strip()
                        newDPT['rw'] = data["rw"]
                        newDPT['rt'] = data["rt"]
                        newDPT['ket'] = data["ket"]
                        results.append(newDPT)
                        no += 1

            
    except Exception as e:
        haveError = True
        print("Error "+filename)
        print(traceback.format_exc())
        # or
        print(sys.exc_info()[2])
        createTxtLog(resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error",filename,traceback.format_exc())

        if not os.path.exists(resultsDir+'/'+dpt["provinsi"]):
            os.makedirs(resultsDir+'/'+dpt["provinsi"])
        if not os.path.exists(resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]):
            os.makedirs(resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"])
        if not os.path.exists(resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error"):
            os.makedirs(resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error")
        shutil.copy(path, resultsDir+'/'+dpt["provinsi"]+"/"+dpt["kabupaten_kota"]+"/error/"+filename)
    
    if len(results) > 0:
        print(len(results),filename)
        saveToCsv(results, str(firstNo) +"_"+str((no-1))+ "_"+ filename,dpt["provinsi"]+"/"+dpt["kabupaten_kota"])

        if deleteOriginal and not haveError:
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
            print(file)
            extraxted = extractData(path+"/"+file,no,dpt)

            # sleep random from 1-4 to make sure the server not block the request
            time.sleep(random.randint(1,4))
    
            no = extraxted["no"]
        else:
            no = deepSearch(path+"/"+file,no,dpt)
    return no


# get folder list
folderList = os.listdir(pdfSourceDir)
dpts = []
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
        "alamat":"-",
        "nomor_tps":1,
        "kelurahan_desa":"KELURAHAN",
        "kecamatan":"KECAMATAN",
        "kabupaten_kota":"KABUPATEN",
        "provinsi":"PROVINSI",
    }
    if os.path.isdir(pdfSourceDir+"/"+folderProvinsi):
        folderKabupatenKotaList = os.listdir(pdfSourceDir+"/"+folderProvinsi)
        dpt["provinsi"] = folderProvinsi
        for folderKabKota in folderKabupatenKotaList:
            
            # remove "SALINAN DPT"
            kabupatenKota = folderKabKota.replace("SALINAN DPT","").replace("_"," ").strip()
            dpt["kabupaten_kota"] = kabupatenKota
            
            no = deepSearch(pdfSourceDir+"/"+folderProvinsi+"/"+folderKabKota,no,dpt)
            print("Done "+kabupatenKota)
        print("Done "+folderProvinsi)

