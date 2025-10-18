import re, fitz
from PIL import Image
import io
import os
from pypdf import PdfReader
from openpyxl import Workbook
import requests
from pathlib import Path

pdfFileName = [
    '03091011',
]
compliance = 'Compliance'
endCompliance = 'Compliance Statements'
productHeader = ['Part Number', 'Product Description', 'Series Number', 'Status', 'Product Category', 'Engineering Number', 'Packaging Alternative']
complianceHeader = ['GADSL/IMDS', 'China RoHS', 'EU ELV', 'Low-Halogen Status', 'REACH SVHC', 'EU RoHS']
physicalHeader = ['Breakaway', 'Circuits (Loaded)','Circuits (maximum)', 'Durability (mating cycles max)', 'Color - Resin', 'Flammability', 'Gender', 
            'Glow-Wire Capable', 'Guide to Mating Part', 'Keying to Mating Part', 'Lock to Mating Part', 'Material - Metal', 'Material - Resin',
            'Material - Plating Mating', 'Material - Plating Termination', 'Plating min - Mating', 'Plating min - Termination', 
            'First Mate / Last Break', 'Net Weight', 'Number of Rows', 'Packaging Type', 'Panel Mount', 'Termination Interface Style', 
            'Pitch - Mating Interfac', 'Pitch - Termination Interface', 'Polarized to Mating Part', 'Stackable', 'Orientation',
            'PCB Locator', 'PCB Retention', 'PCB Thickness - Recommended', 'PC Tail Length', 'Shrouded', 'Wire Insulation Diameter',
            'Temperature Range - Operating', 'Termination Interface Style', 'Wire Size (AWG)', ' Wire Size mm²']
agencyHeader = ['CSA', 'UL', 'Current - Maximum per Contact', 'Voltage - Maximum']
currentPDFDir = f'{os.getcwd()}/molexProducts/pdf'
currentImgDir = f'{os.getcwd()}/molexProducts/img'

def readPDFFile() :
    excelData = []
    currentPDF = Path(currentPDFDir)
    
    for index, name in enumerate(pdfFileName):
        matesResult = []
        useResult = []
        globalResult = []

        pdfFilePath = currentPDF / f"{name}.pdf"

        if not os.path.exists(pdfFilePath):
            print(f'[{index}] {name}.pdf 파일을 찾을 수 없습니다. 건너뜁니다.')
            continue

        try :
            reader = PdfReader(pdfFilePath)
            fullText = ""
            print(f'---------------------------- {index + 1}번째 {name} 파일을 읽기 시작합니다. ------------------------------------------- ')
            for page in range(len(reader.pages)):
                text = reader.pages[page].extract_text()
                fullText += text + '\n'

                if int(page) == 0 :
                    product_info = makeDefaultDataExcel(text)
            
            image = uploadUrlImage(name)
            complianceResult = makeComplianceData(fullText)
            physical_data = makePhysicalData(fullText)
            agencyResult = makeAgencyData(fullText)
            matesResult.append(makeMatesWithPartData(fullText))
            useResult.append(makeUseWithPartData(fullText))
            globalResult.append(makeGlobalData(fullText))

            finalResult['image'] = image
            finalResult = {**product_info, **(complianceResult or {}), **(physical_data or {}), **(agencyResult or {})}
            finalResult['mateswith'] = matesResult
            finalResult['useWith'] = useResult
            finalResult['global'] = globalResult

            print(finalResult)
            # excelData.append(final_data)

            print(f'---------------------------- {index + 1}번째 {name} 파일을 완료하였습니다. ------------------------------------------- ')
            print('\n')
            
        except Exception as e:
            print(f'[{index}] {name} 파일 처리 중 오류 발생: {e}')
            continue

    return excelData

def makeDefaultDataExcel(text):
    if any(header in text for header in productHeader):
        data = {}
        for headers in productHeader:
            pattern = rf'{re.escape(headers)}\s*:\s*(.+)'
            match = re.search(pattern, text)
            if match:
                header = changeHeader(headers)
                data[header] = match.group(1).strip()

        return data
    return None


def makeComplianceData(fullText):
    match = re.search(r'Compliance(.*?)Compliance Statements', fullText, re.DOTALL | re.IGNORECASE)
    data = {}
    result = {}

    if match:
        extracted = match.group(1).strip()
        result['compliance'] = makeJsonData(extracted, data, complianceHeader)
        return result

def makePhysicalData(fullText):
    match = re.search(r'Physical(.*?)Mates With / Use With', fullText, re.DOTALL | re.IGNORECASE)
    data = {}
    result = {}

    if match:
        extracted = match.group(1).strip()
        result['physical'] = makeJsonData(extracted, data, physicalHeader)
        return result

def makeAgencyData(fullText):
    match = re.search(r'Agency(.*?)Physical', fullText, re.DOTALL | re.IGNORECASE)
    data = {}
    result = {}

    if match:
        extracted = match.group(1).strip()
        result['agency'] = makeJsonData(extracted, data, agencyHeader)
        return result

def makeMatesWithPartData(fullText):
    match = re.search(r'Mates with Part\(s\)(.*?)Use with Part\(s\)', fullText, re.DOTALL | re.IGNORECASE)
    if match:
        result = makeDescPartNo(match)
        return result
        

def makeUseWithPartData(fullText):
    match = re.search(r'Use with Part\(s\)(.*?)(?:Application Tooling|$)', fullText, re.DOTALL | re.IGNORECASE)
    if match:
        result = makeDescPartNo(match)
        return result

def makeGlobalData(fullText):
    match = re.search(r'Global(.*?)Japan', fullText, re.DOTALL | re.IGNORECASE)
    if match:
        result = makeDescPartNo(match)
        return result


def makeJsonData(extracted, data, header_list):
    cleaned = re.sub(r'\s+', ' ', extracted)
    sliceHeader = '|'.join([re.escape(h) for h in header_list])
    pattern = rf'({sliceHeader})\s+(.*?)(?={sliceHeader}|$)'

    for m in re.finditer(pattern, cleaned):
        keys = m.group(1).strip()
        value = m.group(2).strip()

        key = changeHeader(keys)
        data[key] = value
    return data

def makeDescPartNo(match):
    result = {}
    extracted = match.group(1).strip()

    cleaned = re.sub(r'\s+', ' ', extracted)
    removeCleaned = cleaned.replace('Description Part Number', '').strip()
    removeCleaned = re.sub(r'This document was generated on[^.]*\.?', '', removeCleaned, flags=re.IGNORECASE).strip()
    parts = re.split(r'(\d{4,})', removeCleaned)
    for i in range(0, len(parts)-1, 2):
        desc = parts[i].strip()
        partNo = parts[i+1].strip() if i+1 < len(parts) else ''
        if desc and partNo:
            result[partNo] = desc

    return result

def changeHeader(text):
    return text.lower().replace(' ', '')


# def saveToExcel(data_list, filename='output.xlsx'):
#     wb = Workbook()
#     ws = wb.active
#     ws.title = 'Product Data'
    
#     special_headers = [
#         'Mates With PartNo', 'Mates With Desc', 
#         'Use With PartNo', 'Use With Desc', 
#         'Global PartNo', 'Global Desc'
#     ]
    
#     all_headers = (
#         ['NO'] +
#         productHeader + 
#         complianceHeader + 
#         physicalHeader + 
#         agencyHeader +
#         special_headers 
#     )
    
#     final_headers = list(dict.fromkeys(all_headers))
#     ws.append(final_headers)
    
#     for file_index, data in enumerate(data_list):
        
#         mates_list = data.get('Mates With') or [] 
#         use_list = data.get('Use With') or []
#         global_list = data.get('Global') or []

#         max_len = max(len(mates_list), len(use_list), len(global_list))
#         num_rows = max(max_len, 1)

#         for i in range(num_rows):
#             row = []
#             for header in final_headers:
                
#                 if header.startswith('Mates With'):
#                     target_list = mates_list
#                     key = 'partNo' if 'PartNo' in header else 'desc'
#                 elif header.startswith('Use With'):
#                     target_list = use_list
#                     key = 'partNo' if 'PartNo' in header else 'desc'
#                 elif header.startswith('Global'):
#                     target_list = global_list
#                     key = 'partNo' if 'PartNo' in header else 'desc'
#                 else:
#                     target_list = None
                        
#                 if target_list is not None:
#                     if i < len(target_list):
#                         item = target_list[i]
#                         row.append(item.get(key, ''))
#                     else:
#                         row.append('')

#                 else:
#                     value = data.get(header)
                    
#                     if header == 'NO':
#                         row.append(file_index + 1) 
#                     elif value is not None:
#                         row.append(str(value))
#                     else:
#                         row.append('')

#             ws.append(row)
        
#     wb.save(filename)
#     print(f'\n엑셀 파일이 저장되었습니다: {filename}')

# def getImage():
#     currentPDF = Path(currentPDFDir)
#     currentImg = Path(currentImgDir)

#     for index, name in enumerate(pdfFileName):
#         pdfFilePath = currentPDF / f"{name}.pdf"

#         if not os.path.exists(pdfFilePath):
#             print(f'[{index}]doc {name}.pdf 파일을 찾을 수 없습니다. 건너뜁니다.')
#             continue

#         doc = fitz.open(pdfFilePath)
#         for pageNum in range(len(doc)):
#             page = doc[pageNum]
#             images = page.get_images(full=True)

#             for index1, img in enumerate(images):
#                 xref = img[0]
#                 base_image = doc.extract_image(xref)
#                 image_bytes = base_image["image"]
#                 image_ext = base_image["ext"]
#                 imageFilename = currentImg / f"{name}.{image_ext}"

#                 if pageNum == 0 and index1 == 1:
#                     try: 
#                         resultImage = Image.open(io.BytesIO(image_bytes))
#                         resultImage.save(imageFilename)
#                         print(f"Saved: {imageFilename}")

#                     except Exception as e:
#                             print(f"오류 발생, {e}")

#     doc.close()

def uploadUrlImage(name) :
    # url = "https://api.imgbb.com/1/upload"
    # possible = ['.png', '.jpg', '.jpeg']

    # for ext in possible:
    #     imagePath = 'C:/Users/개발팀/OneDrive/Desktop/molexProducts/img'
    #     tempPath = f'{imagePath}/{name}{ext}'
    #     if os.path.exists(tempPath):
    #         imagePath = tempPath
    #         break

    # with open(imagePath, 'rb') as f :
    #     files = {'image': f}
    #     data = {'key' : '64bbef1938be158b5d8c6c6f5f58d3ce'}
    #     response = requests.post(url, files=files, data=data)

    try: 
        # if response.status_code == 200:
        #     result = response.json()
        #     imageUrl = result['data']['url']
        #     return imageUrl
        # else:
        #     return None
        
        return "1"
    except Exception as e:
        print('실행 중 오류 발생', e)

excelData = readPDFFile()
# saveToExcel(excelData, 'product_data.xlsx')

# getImage()
