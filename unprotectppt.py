# unprotect POWERPOINT file v1
# pyinstaller --clean --onefile unprotectppt.py --name unprotectppt.exe --icon unprotectppt_1.ico

import os
import re
import zipfile
import tempfile


# Funkcia vyhlada vsetky excel subory a vrati pole z nazvamy
def find_ppts():
    ls = os.listdir()                       # list dir
    ppts = []
    
    for file in ls:                         # prejde cely dir
        if file.split('.')[-1] == 'pptx':   # posledny prvok oddeleny ciarkou musi byt excel format
            ppts.append(file)               # vlozi do pola nazov najdeneneho excel subora
            
    return ppts


# Funkcia vyhlada zamok a vymaze ho.
def unprotect_powerpoint(file):

    patterns = [
        r'<p:modifyVerifier[^>]+>'
    ]
    
    for pattern in patterns:                    # prejde vsetky paterny
        tags = re.findall(pattern, file)        # vyhlada patern v subore pomocou modulu 're'
        if len(tags) == 1:
            #print(' [unlock presentation] {}'.format(tags[0]))
            break                               # zastavy for ak najde patern
    if len(tags) == 0:                      
        return file                             # vrati neupraveny subor ak nenajde patern
    
    file = file.replace(tags[0], '')            # vymera najdeny patern
    
    return file                                 # vrati upraveny subor


# Vstup je excel subor
def core(ppt):

    path = "ppt/presentation.xml"                                   # cesta pre excel sheety

    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(ppt))   # vytvory tmp subor
    os.close(tmpfd)

    with zipfile.ZipFile(ppt, 'r') as sheet:                      # rozbali excel do zip archivu pre citanie
        with zipfile.ZipFile(tmpname, 'w') as sheetunprotect:       # rozbali tmp excel do zip archivu pre zapis

            for item in sheet.infolist():                           # prejde vsetky subory v origin zip excel
                
                if path not in item.filename:                       # ak subor nie je sheet a ani workbook
                    sheetunprotect.writestr(item, sheet.read(item.filename)) # subor iba zapise do zip archivu
                    
                if path in item.filename:                           # ak najde subor sheet
                    #print(item.filename)
                    data = sheet.read(item.filename)                # precitane data ulozi to premenej
                    data = data.decode('UTF-8')                     # prekodovanie dat do utf-8 pre specialne znaky
                    unprotect = unprotect_powerpoint(data)          # funkcia pre vymazanie zamku, vrati upraveny subor
                    sheetunprotect.writestr(item, unprotect)        # zapise upraveny subor do zip archivu

    new_file = ppt.split('.')                                     # nazov subora oddeli od bodiek, vytvory pole 
    new_file = new_file[:-1] + ['unprotect'] + new_file[-1:]        # posklada novy nazov, stale je to pole
    new_file = '.'.join(new_file)                                   # pospaja pole bodkamy, teraz je to string
    
    try:
        os.rename(tmpname, new_file)                                # modul os, premenuje subor tmp subor (zip archove) za novy
        print('[+] Save to {}'.format(new_file))                
    except:
        os.remove(tmpname)                                          # vymaze tmp subor
        print('[-] already unprotect!')
    print('[+] ...............{} OK'.format('.'*len(ppt)))


if __name__ == '__main__':

    print('\n UNPROTECT POWERPOINT \n')
    print(' Unprotect locked powerpoint presentation.\n')
    
    ppts = find_ppts()                                          # funkcia najde vsetke excel subory 

    print('[+] Najdene subory {}'.format(len(ppts)))
    if len(ppts) > 0:
        vstup = input('\nOdomknut vsetky najdene subory? y/n (default = y) : ')

        if (vstup == 'y' or vstup == '' or vstup == 'Y'):
            for ppt in ppts:                                    # prejde vsetke excel subory
                print('[+] UNPROTECT {}.....'.format(ppt))
                core(ppt)                                         # spusti hlavnu funkciu
        else:
            pass

    input('\n\n Stlac ENTER pre ukoncenie.')
    
