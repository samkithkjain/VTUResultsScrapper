import pytesseract,requests
from PIL import Image
from io import BytesIO
from bs4 import BeautifulSoup
import xlsxwriter

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

url0='http://results.vtu.ac.in/resultsvitavicbcs_19/index.php'
url1='http://results.vtu.ac.in/resultsvitavicbcs_19/captcha_new.php'
url2='http://results.vtu.ac.in/resultsvitavicbcs_19/resultpage.php'

h1={'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
'Accept-Encoding':'gzip, deflate',
'Accept-Language':'en-GB,en-US;q=0.9,en;q=0.8',
'Cache-Control':'max-age=0',
'Connection':'keep-alive',
'Cookie':'PHPSESSID=oc7rrl16qp9ttngqrlhmmjj6k5',
'Host':'results.vtu.ac.in',
'Upgrade-Insecure-Requests':'1',
'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36'}

h2={'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
'Accept-Encoding':'gzip, deflate',
'Accept-Language':'en-GB,en-US;q=0.9,en;q=0.8',
'Cache-Control':'max-age=0',
'Connection':'keep-alive',
'Content-Length':'267',
'Content-Type':'application/x-www-form-urlencoded',
'Cookie':'PHPSESSID=oc7rrl16qp9ttngqrlhmmjj6k5',
'Host':'results.vtu.ac.in',
'Origin': 'http://results.vtu.ac.in',
'Referer':'http://results.vtu.ac.in/resultsvitavicbcs_19/index.php',
'Upgrade-Insecure-Requests': '1',
'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36'}

pay2={'lns':'1BI16CS131',
'captchacode':'52984',
'token':'K2EralVnYUxmS2JxQ2FGK3ZFR28waTdiQWFSLytKa1N6T1JRYkJ3RjBmaXA4LzJBNGVCQkhGblB2RkkwNW4vUmFqZ0l2Y0Q1L1VYZGFtczNQSlhBbXc9PTo66SxcDrL407BZCmws+lsVlA==',
'current_url':'http://results.vtu.ac.in/resultsvitavicbcs_19/index.php'}

def reset():
    r0 = requests.get(url0)
    sessionid=r0.headers['Set-Cookie'].rstrip('; path=/')
    h1['Cookie']=sessionid
    h2['Cookie']=sessionid

    soup = BeautifulSoup(r0.text,'html.parser')
    token=soup.find_all('input',{'value':True})[0].get('value')

    pay2['token']=token

    r1 = requests.get(url1, headers=h1)
    im = Image.open(BytesIO(r1.content))
    im.save(r'C:\Users\SAMKIT~1\AppData\Local\Temp\k.jpg')
    cap=pytesseract.image_to_string(Image.open(r'C:\Users\SAMKIT~1\AppData\Local\Temp\k.jpg'))
    pay2['captchacode']=cap

def grade(marks,pf):
    if pf!='P':
        fail=True
        return 0
    marks=int(marks)
    if marks>=90 and marks<=100:
        return 10
    elif marks>=80 and marks<90:
        return 9
    elif marks>=70 and marks<80:
        return 8
    elif marks>=60 and marks<70:
        return 7
    elif marks>=50 and marks<60:
        return 6
    elif marks>=45 and marks<50:
        return 5
    else:
        return 4

def main():
    max_cp=240
    cp=0
    fail=False
    
    data=xlsxwriter.Workbook("data.xlsx")
    sheet1=data.add_worksheet()
    sheet1.write(0,0,"USN")
    sheet1.write(0,1,"NAME")
    sheet1.write(0,2,"15CS71")
    sheet1.write(0,6,"15CS72")
    sheet1.write(0,10,"15CS73")
    sheet1.write(0,14,"ELE-1")
    sheet1.write(0,19,"ELE-2")
    sheet1.write(0,24,"15CSL76")
    sheet1.write(0,28,"15CSL77")
    sheet1.write(0,32,"15CSP78")
    sheet1.write(0,34,"SGPA")    
    reset()    
    fh=open('usn.txt','r')
    r=1
    for usn in fh:
        try:
            fail=False
            cp=0
            usn=usn.strip()
            pay2['lns']=str(usn)
            r2 = requests.post(url2,headers=h2,data=pay2)
            while "Invalid captcha code !!!" in r2.text:
                print(usn)
                reset()
                r2 = requests.post(url2,headers=h2,data=pay2)
            if "Redirecting to VTU Results Site" in r2.text:
                reset()
                r2 = requests.post(url2,headers=h2,data=pay2)
            soup = BeautifulSoup(r2.text,'html.parser')
            usn=soup.find_all('td')[1].text.lstrip(' : ')
            name=soup.find_all('td')[3].text.lstrip(' : ')
            sem=soup.find_all('b')[6].text[-1]
            print('NAME : %s\tUSN : %s\tSEM : %s'%(name,usn,sem))
            sheet1.write(r,0,usn)
            sheet1.write(r,1,name)
            odiv=soup.find_all('div',{'class':'divTableRow'})
            for s in range(1,9):
                idiv=odiv[s].find_all('div')
                sub=str(idiv[0].text)+'\t'+str(idiv[4].text)+'\t'+str(idiv[5].text)
                print(sub)
                if str(idiv[0].text) == '15CS71':
                    sheet1.write(r,2,str(idiv[2].text))
                    sheet1.write(r,3,str(idiv[3].text))
                    sheet1.write(r,4,str(idiv[4].text))
                    gr=grade(str(idiv[4].text),str(idiv[5].text))
                    cp+=gr*4
                    sheet1.write(r,5,gr)
                elif str(idiv[0].text) == '15CS72':
                    sheet1.write(r,6,str(idiv[2].text))
                    sheet1.write(r,7,str(idiv[3].text))
                    sheet1.write(r,8,str(idiv[4].text))
                    gr=grade(str(idiv[4].text),str(idiv[5].text))
                    cp+=gr*4
                    sheet1.write(r,9,gr)
                elif str(idiv[0].text) == '15CS73':
                    sheet1.write(r,10,str(idiv[2].text))
                    sheet1.write(r,11,str(idiv[3].text))
                    sheet1.write(r,12,str(idiv[4].text))
                    gr=grade(str(idiv[4].text),str(idiv[5].text))
                    cp+=gr*4
                    sheet1.write(r,13,gr)
                elif str(idiv[0].text).startswith('15') and idiv[0].text[-3]=='7' and idiv[0].text[-2]=='4':
                    sheet1.write(r,14,str(idiv[0].text))
                    sheet1.write(r,15,str(idiv[2].text))
                    sheet1.write(r,16,str(idiv[3].text))
                    sheet1.write(r,17,str(idiv[4].text))
                    gr=grade(str(idiv[4].text),str(idiv[5].text))
                    cp+=gr*3
                    sheet1.write(r,18,gr)
                elif str(idiv[0].text).startswith('15') and idiv[0].text[-3]=='7' and idiv[0].text[-2]=='5':
                    sheet1.write(r,19,str(idiv[0].text))
                    sheet1.write(r,20,str(idiv[2].text))
                    sheet1.write(r,21,str(idiv[3].text))
                    sheet1.write(r,22,str(idiv[4].text))
                    gr=grade(str(idiv[4].text),str(idiv[5].text))
                    cp+=gr*3
                    sheet1.write(r,23,gr)
                elif str(idiv[0].text) == '15CSL76':
                    sheet1.write(r,24,str(idiv[2].text))
                    sheet1.write(r,25,str(idiv[3].text))
                    sheet1.write(r,26,str(idiv[4].text))
                    gr=grade(str(idiv[4].text),str(idiv[5].text))
                    cp+=gr*2
                    sheet1.write(r,27,gr)
                elif str(idiv[0].text) == '15CSL77':
                    sheet1.write(r,28,str(idiv[2].text))
                    sheet1.write(r,29,str(idiv[3].text))
                    sheet1.write(r,30,str(idiv[4].text))
                    gr=grade(str(idiv[4].text),str(idiv[5].text))
                    cp+=gr*2
                    sheet1.write(r,31,gr)
                elif str(idiv[0].text) == '15CSP78':
                    sheet1.write(r,32,str(idiv[2].text))
                    gr=grade(str(idiv[4].text),str(idiv[5].text))
                    cp+=gr*2
                    sheet1.write(r,33,gr)
            if not fail:
                sheet1.write(r,34,round(cp/max_cp*10,2))
            r+=1
        except:
            continue
    data.close()
if __name__=="__main__":
    main()
