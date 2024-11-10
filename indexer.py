


'''
###
MISSING Papers:
Schol Accounting 2019 & 2020
Schol Economics 2020
Schol History 2014 & 2018 & 2019 & 2020
###
'''


# importing required modules
import os
import PyPDF2
import PyPDF2.errors
import re
import json
from pathlib import Path
from zipfile import ZipFile
import shutil

indexLol = []
iforgor = {}
subjectList = []

spinner = [
			"( ●    )",
			"(  ●   )",
			"(   ●  )",
			"(    ● )",
			"(     ●)",
			"(    ● )",
			"(   ●  )",
			"(  ●   )",
			"( ●    )",
			"(●     )"
		]


credits = {"One":1,"Two":2,"Three":3,"Four":4,"Five":5,"Six":6,"Seven":7,"Eight":8,"Nine":9,"Ten":10, "Tahi":1, "Rua":2, "Toru":3, "Wh\u0101":4, "Rima":5, "Ono":6, "Whitu":7,"Waru":8,"Iwa":9,"Tekau":10} 

for standardNumber in range(90804,93999):
    os.system("cls")
    print(spinner[standardNumber%9])
    for year in range(2023,2011,-1):

        try:
            # creating a pdf file object
            pdfFileObj = open(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers\exams\\"+str(standardNumber)+"-" + str(year) + ".pdf", 'rb')

            # creating a pdf reader object
            pdfReader = PyPDF2.PdfReader(pdfFileObj)

            # creating a page object
            pageObj = pdfReader.pages[0]
            if re.search(pattern=r"(Level|Kaupae) \d* ([\sA-Za-z()&\-/\u00C0-\u017F]+),* "+str(year)+r"[\S\s]*(Credits|Whiwhinga):", string=pageObj.extract_text()).group(2) not in subjectList:
                subjectList.append(re.search(pattern=r"(Level|Kaupae) \d* ([\sA-Za-z()&\-/\u00C0-\u017F]+),* "+str(year)+r"[\S\s]*(Credits|Whiwhinga):", string=pageObj.extract_text()).group(2))
            
            startYear = 2011

            for testYear in range(2011,2024):
                if Path(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers\exams\\"+str(standardNumber)+"-" + str(testYear) + ".pdf").exists():
                   startYear = testYear
                   break
            indexLol.append({"id": str(len(indexLol)),
                             "title":re.search(pattern=str(standardNumber) + r" {1,}([\sA-Za-z()&/\-\u00C0-\u017F,&]+)(\d[.]\d\d)*[^Q]+(Credits|Whiwhinga)", string=pageObj.extract_text()).group(1),
                               "number":str(standardNumber), "credits": credits[re.search(pattern=r"(Credits|Whiwhinga): ([\w\d]+)", string=pageObj.extract_text()).group(2)],
                                 "subject":re.search(pattern=r"(Level|Kaupae) \d* ([\sA-Za-z()&\-/\u00C0-\u017F]+),* "+str(year)+r"[\S\s]*(Credits|Whiwhinga):", string=pageObj.extract_text()).group(2),
                                   "year-range": str(startYear)+"-"+str(year),
                                     "level":re.search(pattern=r"(Level|Kaupae) (\d*) ([\S\s]+) "+str(year)+r"[\S\s]*(Credits|Whiwhinga):", string=pageObj.extract_text()).group(2)})
            
            # closing the pdf file object
            pdfFileObj.close()

            with ZipFile(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers\zipped\\"+str(standardNumber)+".zip", mode="w") as archive:
                for i in range(startYear,year+1):
                    if Path(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers\exams\\"+str(standardNumber)+"-" + str(i) + ".pdf").exists():
                        archive.write(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers\exams\\"+str(standardNumber)+"-" + str(i) + ".pdf", r"exams\\"+str(standardNumber)+"-" + str(i) + ".pdf")
                    if Path(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers\schedules\\"+str(standardNumber)+"-" + str(i) + ".pdf").exists():
                        archive.write(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers\schedules\\"+str(standardNumber)+"-" + str(i) + ".pdf", r"schedules\\"+str(standardNumber)+"-" + str(i) + ".pdf")
            break

        except Exception as e:
            print(e)
            if type(e) is not FileNotFoundError:
                iforgor[str(standardNumber)] = str(e)

for filename in os.listdir(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers/exams/bonus/"):
    f = os.path.join(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers/exams/bonus/", filename)

    # checking if it is a file
    split = filename.split('.')
    if os.path.isfile(f) and split[6] == "zip":
        subjectList.append(split[3])
        indexLol.append({"id": str(len(indexLol)),
                             "title":split[0],
                               "number":split[1], "credits": split[2],
                                 "subject":split[3],
                                   "year-range": split[4],
                                     "level":split[5]})
        shutil.copyfile(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers/exams/bonus/"+filename,r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers/zipped/" +split[1]+".zip")

with open(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers\website\searchIndex.json", "w") as outfile:
    json.dump(indexLol, outfile)

with open(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers\website\brokenFiles.json", "w") as outfile:
    json.dump(iforgor, outfile)

with open(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers\website\subjects.json", "w") as outfile:
    subjectList.sort()
    json.dump(subjectList, outfile)