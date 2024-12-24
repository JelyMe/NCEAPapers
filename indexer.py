


'''
###
MISSING Papers:
Schol Accounting 2019 & 2020
Schol Economics 2020
Schol History 2014 & 2018 & 2019 & 2020
Schol English 2018 & 2019 & 2020
Schol Art History 2014 & 2018 & 2019 & 2020
Schol Classics 2018 & 2019 & 2020
Schol Media Studies 2018 & 2019 & 2020
Schol Chinese 2018 & 2019 & 2020
Schol French 2015 & 2016 & 2018 & 2019 & 2020
Schol German 2014 & 2018 & 2019 & 2020
Schol Latin 2014 & 2016 & 2018 & 2019 & 2020
Schol Japanese 2014 & 2018 & 2019 & 2020
###
'''

r'''
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

for standardNumber in range(30000,93999):
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
                    if Path(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers\resources\\"+str(standardNumber)+"-" + str(i) + ".pdf").exists():
                        archive.write(r"C:\Users\jelym\Documents\NZQA Papers\NCEAPapers\resources\\"+str(standardNumber)+"-" + str(i) + ".pdf", r"resources\\"+str(standardNumber)+"-" + str(i) + ".pdf")
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
'''

from multiprocessing.pool import ThreadPool
import os
import PyPDF2
import PyPDF2.errors
import re
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import shutil
import tqdm
import webbrowser
import pathlib

paperList = {}
subjectList = []

credits = {"One":1,"Two":2,"Three":3,"Four":4,"Five":5,"Six":6,"Seven":7,"Eight":8,"Nine":9,"Ten":10, "Tahi":1, "Rua":2, "Toru":3, "Wh\u0101":4, "Rima":5, "Ono":6, "Whitu":7,"Waru":8,"Iwa":9,"Tekau":10} 

def indexPaper(filename):
    delete = False
    f = os.path.join("./exams/", filename)
    with open(f, 'rb') as file:
        if file.read(2) != b'%P':
            delete = True
    if not delete:
        fileInfo = re.search(r'(\w+)-(\w+)-?(\w+)?\.pdf', filename)
        if fileInfo != None:
            if fileInfo[3] == None:
                if paperList.get(fileInfo[1]) == None:
                    paperList[fileInfo[1]] = {
                        "number":fileInfo[1], 
                        "start-year":fileInfo[2], 
                        "end-year":fileInfo[2]
                    }
                    try:
                        # creating a pdf file object
                        pdfFileObj = open(f, 'rb')

                        # creating a pdf reader object
                        pdfReader = PyPDF2.PdfReader(pdfFileObj)

                        # creating a page object
                        pageObj = pdfReader.pages[0]
                        pageText = pageObj.extract_text()
                        if re.search(pattern=r"(Level|Kaupae) \d* ([\sA-Za-z()&\-/\u00C0-\u017F,]+),* "+fileInfo[2]+r"[\S\s]*(Credits|W?hiwhinga):", string=pageText).group(2) not in subjectList:
                            subjectList.append(re.search(pattern=r"(Level|Kaupae) \d* ([\sA-Za-z()&\-/\u00C0-\u017F,]+),* "+fileInfo[2]+r"[\S\s]*(Credits|W?hiwhinga):", string=pageText).group(2))
                        
                        endYear = fileInfo[2]
                        if paperList[fileInfo[1]]["end-year"] != fileInfo[2]:
                            endYear = paperList[fileInfo[1]]["end-year"]

                        paperList[fileInfo[1]] = {
                            "number":fileInfo[1], 
                            "start-year":fileInfo[2], 
                            "end-year":fileInfo[2], 
                            "title":re.search(pattern=fileInfo[1] + r" {1,}([\sA-Za-z()&/\-\u00C0-\u017F,&]+)(\d[.]\d\d)*[^Q]+(Credits|W?hiwhinga)", string=pageText).group(1),
                            "number":fileInfo[1], 
                            "credits": credits[re.search(pattern=r"(Credits|W?hiwhinga): ([\w\d]+)", string=pageText).group(2)],
                            "subject":re.search(pattern=r"(Level|Kaupae) \d* ([\sA-Za-z()&\-/\u00C0-\u017F,]+),* "+fileInfo[2]+r"[\S\s]*(Credits|W?hiwhinga):", string=pageText).group(2),
                            "level":re.search(pattern=r"(Level|Kaupae) (\d*) ([\S\s]+) "+fileInfo[2]+r"[\S\s]*(Credits|W?hiwhinga):", string=pageText).group(2)
                        }

                        paperList[fileInfo[1]]["end-year"] = endYear
                        
                        # closing the pdf file object
                        pdfFileObj.close()
                    except Exception as e:
                        pass
                else:
                    paperList[fileInfo[1]]["end-year"] = fileInfo[2]
                paperList[fileInfo[1]]["id"] = next((i for i, d in enumerate(paperList.values()) if d.get("number") == fileInfo[1]), None)
            else:
                if paperList.get(fileInfo[1]) == None:
                    paperList[fileInfo[1]] = {
                        "number":fileInfo[1], 
                        "start-year":fileInfo[2], 
                        "end-year":fileInfo[2]
                    }

                    # creating a pdf file object
                    pdfFileObj = open(f, 'rb')

                    # creating a pdf reader object
                    pdfReader = PyPDF2.PdfReader(pdfFileObj)

                    # creating a page object
                    pageObj = pdfReader.pages[0]
                    pageText = pageObj.extract_text()
                    if re.search(pattern=r"([\sA-Za-z()&\-/\u00C0-\u017F,]+),* "+fileInfo[2]+r"[\S\s]*(Credits|W?hiwhinga):", string=pageText).group(2) not in subjectList:
                        subjectList.append(re.search(pattern=r"([\sA-Za-z()&\-/\u00C0-\u017F,]+),* "+fileInfo[2]+r"[\S\s]*(Credits|W?hiwhinga):", string=pageText).group(2))
                    
                    endYear = fileInfo[2]
                    if paperList[fileInfo[1]]["end-year"] != fileInfo[2]:
                        endYear = paperList[fileInfo[1]]["end-year"]

                    paperList[fileInfo[1]] = {
                        "number":fileInfo[1], 
                        "start-year":fileInfo[2], 
                        "end-year":fileInfo[2], 
                        "title":re.search(pattern=fileInfo[1] + r" {1,}([\sA-Za-z()&/\-\u00C0-\u017F,&]+)(\d[.]\d\d)*[^Q]+(Credits|W?hiwhinga)", string=pageText).group(1),
                        "number":fileInfo[1], 
                        "credits": credits[re.search(pattern=r"(Credits|W?hiwhinga): ([\w\d]+)", string=pageText).group(2)],
                        "subject":re.search(pattern=r"([\sA-Za-z()&\-/\u00C0-\u017F,]+),* "+fileInfo[2]+r"[\S\s]*(Credits|W?hiwhinga):", string=pageText).group(2),
                        "level":"All"
                    }

                    paperList[fileInfo[1]]["end-year"] = endYear
                    
                    # closing the pdf file object
                    pdfFileObj.close()
                else:
                    paperList[fileInfo[1]]["end-year"] = fileInfo[2]
                paperList[fileInfo[1]]["id"] = next((i for i, d in enumerate(paperList.values()) if d.get("number") == fileInfo[1]), None)

        else:
            os.remove(f)

index = None
if input("Skip Index Y/N: ").lower() != "y":
    with ThreadPool() as pool:
        files = [f for f in os.listdir("./exams/") if os.path.isfile(os.path.join("./exams/", f))]
        for _ in tqdm.tqdm(pool.imap(indexPaper, files), total=len(files), desc="Indexing Papers", unit="files"):
            pass
    
    index = list(paperList.values())
    with open("searchIndex.json", "w") as outfile:
        json.dump(index, outfile, indent=2)
    with open("subjects.json", "w") as outfile:
        json.dump(subjectList, outfile, indent=2)
else:
    with open("searchIndex.json", "r") as file:
        index = json.load(file)
    with open("subjects.json", "r") as file:
        subjectList = json.load(file)

untitled = [d for d in index if 'title' not in d]

override = {}
with open("manualOverride.json", "r") as file:
        override = json.load(file)

def fixFiles(info):
    if info["number"] in override:
        index[info["id"]] = override[info["number"]]
        index[info["id"]]["id"] = info["id"]
        index[info["id"]]["number"] = info["number"]
        index[info["id"]]["start-year"] = info["start-year"]
        index[info["id"]]["end-year"] = info["end-year"]
        if override[info["number"]]["subject"] not in subjectList:
            subjectList.append(override[info["number"]]["subject"])
    else:
        f = os.path.join("./exams/", info["number"]+"-"+info["end-year"]+".pdf")
        if index[info["id"]].get("title") == None:
            try:
                # creating a pdf file object
                pdfFileObj = open(f, 'rb')

                # creating a pdf reader object
                pdfReader = PyPDF2.PdfReader(pdfFileObj)

                # creating a page object
                pageObj = pdfReader.pages[0]
                pageText = pageObj.extract_text()
                index[info["id"]]["title"] = re.search(pattern=info["number"] + r" {1,}([\sA-Za-z()&/\-\u00C0-\u017F,&]+)(\d[.]\d\d)*[^Q]+(Credits|W?hiwhinga)", string=pageText).group(1)
                index[info["id"]]["credits"] = credits[re.search(pattern=r"(Credits|W?hiwhinga): ([\w\d]+)", string=pageText).group(2)]
                index[info["id"]]["subject"] = re.search(pattern=r"(Level|Kaupae) \d* ([\sA-Za-z()&\-/\u00C0-\u017F,]+),* "+info["end-year"]+r"[\S\s]*(Credits|W?hiwhinga):", string=pageText).group(2)
                index[info["id"]]["level"] = re.search(pattern=r"(Level|Kaupae) (\d*) ([\S\s]+) "+info["end-year"]+r"[\S\s]*(Credits|W?hiwhinga):", string=pageText).group(2)

                # closing the pdf file object
                pdfFileObj.close()
            except Exception as e:
                print(info["number"] + " - " + str(e))

with ThreadPool() as pool:
    for _ in tqdm.tqdm(pool.imap(fixFiles, untitled), total=len(untitled), desc="Attempting Automatic Fixes", unit="files"):
        pass

untitled = [d for d in index if 'title' not in d]

if len(untitled) > 0:
    if input("There are "+ str(len(untitled)) +" file(s) without indexes, would you like to manually add them? Y/N: ").lower() == "y":
        for info in untitled:
            webbrowser.open("file://"+str(pathlib.Path().resolve())+"/exams/"+info["number"]+"-"+info["end-year"]+".pdf")
            override[info["number"]] = {}
            title = input("Title: ")
            index[info["id"]]["title"] = title
            override[info["number"]]["title"] = title
            credits = input("Credits: ")
            index[info["id"]]["credits"] = credits
            override[info["number"]]["credits"] = credits
            subject = input("Subject: ")
            index[info["id"]]["subject"] = subject
            override[info["number"]]["subject"] = subject
            if subject not in subjectList:
                subjectList.append(subject)
            level = input("Level: ")
            index[info["id"]]["level"] = level
            override[info["number"]]["level"] = level

with open("searchIndex.json", "w") as outfile:
    json.dump(index, outfile, indent=2)
with open("subjects.json", "w") as outfile:
    json.dump(subjectList, outfile, indent=2)
with open("manualOverride.json", "w") as outfile:
    json.dump(override, outfile, indent=2)

def zipPaper(info):
    if info.get("title") != None:
        with ZipFile("./zipped/"+info["number"]+".zip", mode="w", compression=ZIP_DEFLATED) as archive:
            for i in range(int(info["start-year"]), int(info["end-year"]) + 1):
                if Path("./exams/"+info["number"]+"-" + str(i) + ".pdf").exists():
                    archive.write("./exams/"+info["number"]+"-" + str(i) + ".pdf", "exams/"+info["number"]+"-" + str(i) + ".pdf")
                if Path("./exams/"+info["number"]+"-" + str(i) + "-Term1.pdf").exists():
                    archive.write("./exams/"+info["number"]+"-" + str(i) + "-Term1.pdf", "exams/"+info["number"]+"-" + str(i) + "-Term1.pdf")
                if Path("./exams/"+info["number"]+"-" + str(i) + "-Term2.pdf").exists():
                    archive.write("./exams/"+info["number"]+"-" + str(i) + "-Term2.pdf", "exams/"+info["number"]+"-" + str(i) + "-Term2.pdf")
                if Path("./exams/"+info["number"]+"-" + str(i) + "-Term3.pdf").exists():
                    archive.write("./exams/"+info["number"]+"-" + str(i) + "-Term3.pdf", "exams/"+info["number"]+"-" + str(i) + "-Term3.pdf")
                if Path("./exams/"+info["number"]+"-" + str(i) + "-Term4.pdf").exists():
                    archive.write("./exams/"+info["number"]+"-" + str(i) + "-Term4.pdf", "exams/"+info["number"]+"-" + str(i) + "-Term4.pdf")
                if Path("./schedules/"+info["number"]+"-" + str(i) + ".pdf").exists():
                    archive.write("./schedules/"+info["number"]+"-" + str(i) + ".pdf", "schedules/"+info["number"]+"-" + str(i) + ".pdf")
                if Path("./schedules/"+info["number"]+"-" + str(i) + "-Term1.pdf").exists():
                    archive.write("./schedules/"+info["number"]+"-" + str(i) + "-Term1.pdf", "schedules/"+info["number"]+"-" + str(i) + "-Term1.pdf")
                if Path("./schedules/"+info["number"]+"-" + str(i) + "-Term2.pdf").exists():
                    archive.write("./schedules/"+info["number"]+"-" + str(i) + "-Term2.pdf", "schedules/"+info["number"]+"-" + str(i) + "-Term2.pdf")
                if Path("./schedules/"+info["number"]+"-" + str(i) + "-Term3.pdf").exists():
                    archive.write("./schedules/"+info["number"]+"-" + str(i) + "-Term3.pdf", "schedules/"+info["number"]+"-" + str(i) + "-Term3.pdf")
                if Path("./schedules/"+info["number"]+"-" + str(i) + "-Term4.pdf").exists():
                    archive.write("./schedules/"+info["number"]+"-" + str(i) + "-Term4.pdf", "schedules/"+info["number"]+"-" + str(i) + "-Term4.pdf")
                if Path("./resources/"+info["number"]+"-" + str(i) + ".pdf").exists():
                    archive.write("./resources/"+info["number"]+"-" + str(i) + ".pdf", "resources/"+info["number"]+"-" + str(i) + ".pdf")
                if Path("./audio/"+info["number"]+"-" + str(i) + ".pdf").exists():
                    archive.write("./audio/"+info["number"]+"-" + str(i) + ".pdf", "audio/"+info["number"]+"-" + str(i) + ".pdf")
                if Path("./audio/"+info["number"]+"-" + str(i)).exists():
                    for root, dirs, files in os.walk("./audio/"+info["number"]+"-" + str(i)):
                        for file in files:
                            archive.write(os.path.join(root, file), "audio/"+file)

with ThreadPool() as pool:
    for _ in tqdm.tqdm(pool.imap(zipPaper, index), total=len(index), desc="Zipping Papers", unit="files"):
        pass