import requests
import re
import random
import base64
import json
import os,csv
import pandas as pd
import datetime
from urllib.request import urlretrieve
import openpyxl

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
    urlretrieve(download_info['url_download'], download_info['download_path']+"/"+download_info['file_name'])
    return download_info['download_path']+"/"+download_info['file_name']


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
        fieldnames = ['no', 'nama', 'jenis_kelamin', 'usia', 'rt', 'rw', 'nik', 'ket','alamat', 'nomor_tps', 'kelurahan_desa', 'kecamatan', 'kabupaten_kota', 'provinsi']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for result in results:
            writer.writerow(result)
    csvFile = pd.read_csv(csvFileName, encoding='cp1252')
    xlsxFileName = './results/'+parent+'/excel/'+fileName+'.xlsx'
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
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'accept': 'application/json',
        'authorization': f'Bearer {token}',
        'sec-ch-ua': '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }

    upload_url = f"https://{random_server}.ilovepdf.com/v1/upload"
    response = requests.post(upload_url, headers=headers, files=files, data=payload)
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
            "download_path": "results"
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
    response = convert_excel(path)
    # Define variable to load the dataframe
    dataframe = openpyxl.load_workbook(response['data']['download_path']+"/"+response['data']['file_name'])

    # Define variable to read sheet
    dataframe1 = dataframe.active

    # Iterate the loop to read the cell values
    read = False
    results = []
    tps = 0
    
    for row in range(0, dataframe1.max_row):

        data = []
        for col in dataframe1.iter_cols(1, dataframe1.max_column):
            if col[row].value != None:
                value = str(col[row].value).strip()
                if value == "NAMA":
                    read = True
                    break
                elif "Rekapitulasi" in value:
                    read = False
                elif read and value not in "1 2 3 4 5 6 7 8 9":
                    data.append(col[row].value)
                elif len(value.split(":")) == 4:
                    valueSplit = value.split(":")
                    tps = valueSplit[3]
                    dpt['kelurahan_desa'] = valueSplit[2].strip()
                    dpt['kecamatan'] = valueSplit[1].strip()


        if read and len(data) >= 7:
            newDPT = dpt.copy()
            newDPT['no'] = no
            newDPT['nomor_tps'] =  tps
            newDPT['nama'] = data[1]
            newDPT['jenis_kelamin'] = data[2]
            newDPT['usia'] = data[3]
            newDPT['alamat'] = data[4]
            newDPT['rw'] = data[5]
            newDPT['rt'] = data[6]
            if len(data) > 7:
                newDPT['ket'] = data[7]
            results.append(newDPT)
            no += 1
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

# file_path = './pdfs/A-KabKo-(70593) SEPAKU-TELEMOW_TPS 6.pdf'  # Provide the full path to the local PDF file
# response = convert_excel(file_path)
# print(response)

pdfSourceDir =  "./pdf-sources"

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
            
            no = deepSearch("./pdf-sources/"+folderProvinsi+"/"+folderKabKota,no,dpt)
            print("Done "+kabupatenKota)
        print("Done "+folderProvinsi)

