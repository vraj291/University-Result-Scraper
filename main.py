from config import *
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup 
import csv
import time
import pandas as pd
import pdfkit as pdf
import os
import math
import pickle

def getSGPA(e):
    if e == {}:
        return 0
    return float(e['sgpa'])

def getIds():
    ids = []
    # for i in range(CLASS_SIZE[0],CLASS_SIZE[1]+1):
    for i in range(2,5):
        _id =  D2D_ID_PREFIX if i>NON_D2D_LAST_ID else NON_D2D_ID_PREFIX
        _id += '0' * (2 - (int)(math.log10(i))) + str(i)
        ids.append(_id)
    return ids

def getSubjectsforId(s,g):
    subjects = []
    grdCount = 0
    temp={}
    for sub in s:
        if len(sub.find('span').text.strip()) == 0:
            continue
        if sub.find('span').text.strip()[2].isdigit() :
            subjects.append(temp)
            temp={}
            temp['subject_code'] = sub.find('span').text.strip()
            temp['exam'] = {'theory':{},'practical':{}}
        elif sub.text.strip() == 'THEORY': 
            temp['exam']['theory']['credit']=g[grdCount].find('span').text.strip()
            grdCount +=1
            temp['exam']['theory']['grade']=g[grdCount].find('span').text.strip()
            grdCount +=1
        elif sub.text.strip() == 'PRACTICAL': 
            temp['exam']['practical']['credit']=g[grdCount].find('span').text.strip()
            grdCount +=1
            temp['exam']['practical']['grade']=g[grdCount].find('span').text.strip()
            grdCount +=1
        else:
            temp['subject_name']=sub.find('span').text.strip()
    subjects.append(temp)
    return subjects[1:]

def getDetailsforId(id,driver):

    sel = Select(driver.find_element_by_id("ddlInst"))
    sel.select_by_visible_text(INSTITUTE)
    sel = Select(driver.find_element_by_id("ddlDegree"))
    sel.select_by_visible_text(DEGREE)
    sel = Select(driver.find_element_by_id("ddlSem"))
    sel.select_by_visible_text(SEM)
    sel = Select(driver.find_element_by_id("ddlScheduleExam"))
    sel.select_by_visible_text(EXAM_SCHEDULE)
    driver.find_element_by_id("txtEnrNo").send_keys(id)
    driver.find_element_by_id("btnSearch").send_keys(Keys.ENTER) 

    score = {}

    html = BeautifulSoup(driver.page_source,features="html5lib")
    name = html.find('span',{'id':'uclGrd1_lblStudentName'})
    if name is None:
        return score
    subs = html.find('table',{'id':'uclGrd1_grdResult'}).find_all('td',{'class' : "Newtd"})
    grds = html.find_all('td',{'class' : "GrdCel"})
    sgpa = html.find('span',{'id':'uclGrd1_lblSGPA'})

    score['id'] = id
    score['name'] = name.text.strip()
    score['subjects'] = getSubjectsforId(subs,grds)
    score['sgpa'] = sgpa.text.strip()

    driver.find_element_by_id("btnBack1").send_keys(Keys.ENTER) 

    return score

def getAllDetails(driver) : 
    ids = getIds()
    scores = []
    for id in ids:
        scores.append(getDetailsforId(id,driver))
    scores.sort(reverse=True,key=getSGPA)
    return scores

def createFile(driver) : 
    scores = getAllDetails(driver)
    with open(f'{RESULT_FILE_NAME}.csv', 'w', newline='') as file:
        fieldnames = ['ID','NAME']
        for i in scores[0]['subjects']:
            if i['subject_code'] not in SUBJECT_CODES:
                if len(i['exam']['theory']) != 0:
                    fieldnames.append('ELECTIVE\n(THEORY)') 
                if len(i['exam']['practical']) != 0:
                    fieldnames.append('ELECTIVE\n(PRACTICAL)')
            else:
                if len(i['exam']['theory']) != 0:
                    fieldnames.append(i['subject_code']+'\n(THEORY)') 
                if len(i['exam']['practical']) != 0:
                    fieldnames.append(i['subject_code']+'\n(PRACTICAL)')
        fieldnames.append('SGPA')
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for i in scores:
            if i == {}:
                continue
            res = {'ID': i['id'], 'NAME': i['name'], 'SGPA': i['sgpa']}
            for j in i['subjects']:
                sub_code = j['subject_code']
                if sub_code not in SUBJECT_CODES:
                    sub_code = 'ELECTIVE'
                if len(j['exam']['theory']) != 0:
                    res[sub_code+'\n(THEORY)'] = j['exam']['theory']['grade']+'('+j['exam']['theory']['credit']+')'
                if len(j['exam']['practical']) != 0:
                    res[sub_code+'\n(PRACTICAL)'] = j['exam']['practical']['grade']+'('+j['exam']['practical']['credit']+')'
            writer.writerow(res)

# dbfile = open('examplePickle', 'ab')
# pickle.dump({'results':scores}, dbfile)                     
# dbfile.close()

def getResults() : 
    path = PATH_TO_CHROME_DRIVER 
    page = SOURCE_PAGE
    driver = webdriver.Chrome(path)
    driver.get(page) 
    createFile(driver)
    df = pd.read_csv(f'{RESULT_FILE_NAME}.csv',sep=',')
    df.to_html('results.html')
    config = pdf.configuration(wkhtmltopdf=PATH_TO_WKHTMLTOPDF)
    pdf.from_url("results.html", f'{RESULT_FILE_NAME}.pdf', configuration=config)
    os.remove('results.html')
    driver.close()

if __name__=='__main__' : 
    getResults()