# -*- coding: utf-8 -*-

from flask import Flask
from flask import g
from flask import url_for, redirect, render_template, request, send_from_directory
from flask.ext.wtf import Form
from wtforms import TextField, PasswordField, validators, SelectField, SubmitField
from flask.ext.login import LoginManager, UserMixin, login_user, login_required, current_user
from flask_bootstrap import Bootstrap
import sqlite3
import pickle
import xlwt
import urllib2
import urllib
import cookielib
import re
from pyquery import PyQuery as pq

app = Flask(__name__)
app.config['DEBUG'] = False
app.config['SECRET_KEY'] = 'glassices'

login_manager = LoginManager()
login_manager.init_app(app)
Bootstrap(app)

DATABASE = 'data.db'

def make_dict(cursor, row):
    return dict((cursor.description[idx][0], value) for idx, value in enumerate(row))

def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect(DATABASE)
    return db

@app.teardown_appcontext
def close_connection(exception):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()

def init_the_database():
    try:
        grade = {}

        def decode_grade(x):
            number = x[0:-1]
            if not grade.has_key(number):
                grade[number] = len(grade)
            return grade[number]

        username = "zhaoyuan13"
        password = "LYWVT4Q6"
        hard_code1 = "linkSql=&is_input_orderItem_0=3902%5Easc&is_input_orderItem_1=3929%5Easc&is_input_orderItem_2=3908%5Easc&is_input_orderItem_3=3918%5Easc&is_input_orderItemCount=4&is_input_orderItem_0=3902%5Easc&is_input_orderItem_1=3929%5Easc&is_input_orderItem_2=3908%5Easc&is_input_orderItem_3=3918%5Easc&is_input_orderItemCount=4&is_input_orderItem_0=3902%5Easc&is_input_orderItem_1=3929%5Easc&is_input_orderItem_2=3908%5Easc&is_input_orderItem_3=3918%5Easc&is_input_orderItemCount=4&is_input_orderItem_0=3902%5Easc&is_input_orderItem_1=3929%5Easc&is_input_orderItem_2=3908%5Easc&is_input_orderItem_3=3918%5Easc&is_input_orderItemCount=4&id=243&currentPage="
        hard_code2 = "&yhxzl=&pageCount=13&hid_div_disp=none&tjsize=13&cxl0=10&ifftj0=0&ztjIdStr0=&ztjListStr0=&srchStr0=&cxid0=3902&cxczf0=%3D&3902=&cxl1=20&ifftj1=0&ztjIdStr1=&ztjListStr1=&srchStr1=&cxid1=3903&cxczf1=like&3903=&cxl2=25&ifftj2=1&ztjIdStr2=3907&ztjListStr2=%40%40value%3D%BB%AF%D1%A701%40%40text%3D%BB%AF%D1%A701%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A701%40%40text%3D%BB%AF%D1%A701%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A702%40%40text%3D%BB%AF%D1%A702%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A702%40%40text%3D%BB%AF%D1%A702%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A711%40%40text%3D%BB%AF%D1%A711%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A711%40%40text%3D%BB%AF%D1%A711%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A712%40%40text%3D%BB%AF%D1%A712%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A712%40%40text%3D%BB%AF%D1%A712%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A721%40%40text%3D%BB%AF%D1%A721%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A721%40%40text%3D%BB%AF%D1%A721%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A722%40%40text%3D%BB%AF%D1%A722%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A722%40%40text%3D%BB%AF%D1%A722%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A731%40%40text%3D%BB%AF%D1%A731%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A731%40%40text%3D%BB%AF%D1%A731%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A731%40%40text%3D%BB%AF%D1%A731%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A732%40%40text%3D%BB%AF%D1%A732%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A732%40%40text%3D%BB%AF%D1%A732%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A732%40%40text%3D%BB%AF%D1%A732%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A741%40%40text%3D%BB%AF%D1%A741%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A741%40%40text%3D%BB%AF%D1%A741%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A741%40%40text%3D%BB%AF%D1%A741%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A742%40%40text%3D%BB%AF%D1%A742%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A742%40%40text%3D%BB%AF%D1%A742%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A742%40%40text%3D%BB%AF%D1%A742%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A751%40%40text%3D%BB%AF%D1%A751%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A751%40%40text%3D%BB%AF%D1%A751%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A751%40%40text%3D%BB%AF%D1%A751%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A752%40%40text%3D%BB%AF%D1%A752%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A752%40%40text%3D%BB%AF%D1%A752%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A752%40%40text%3D%BB%AF%D1%A752%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A761%40%40text%3D%BB%AF%D1%A761%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A761%40%40text%3D%BB%AF%D1%A761%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A762%40%40text%3D%BB%AF%D1%A762%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A762%40%40text%3D%BB%AF%D1%A762%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A771%40%40text%3D%BB%AF%D1%A771%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A772%40%40text%3D%BB%AF%D1%A772%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A781%40%40text%3D%BB%AF%D1%A781%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A781%40%40text%3D%BB%AF%D1%A781%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A782%40%40text%3D%BB%AF%D1%A782%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A782%40%40text%3D%BB%AF%D1%A782%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A791%40%40text%3D%BB%AF%D1%A791%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A791%40%40text%3D%BB%AF%D1%A791%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A792%40%40text%3D%BB%AF%D1%A792%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%AF%D1%A792%40%40text%3D%BB%AF%D1%A792%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%F9%BF%C609%40%40text%3D%BB%F9%BF%C609%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%F9%BF%C613%40%40text%3D%BB%F9%BF%C613%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%F9%BF%C623%40%40text%3D%BB%F9%BF%C623%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%F9%BF%C633%40%40text%3D%BB%F9%BF%C633%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%F9%BF%C633%40%40text%3D%BB%F9%BF%C633%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%F9%BF%C643%40%40text%3D%BB%F9%BF%C643%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%F9%BF%C643%40%40text%3D%BB%F9%BF%C643%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%F9%BF%C653%40%40text%3D%BB%F9%BF%C653%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%F9%BF%C659%40%40text%3D%BB%F9%BF%C659%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%F9%BF%C669%40%40text%3D%BB%F9%BF%C669%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%F9%BF%C679%40%40text%3D%BB%F9%BF%C679%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%F9%BF%C689%40%40text%3D%BB%F9%BF%C689%40%40tjvalue%3D044%2C%23%2C%23%40%40value%3D%BB%F9%BF%C699%40%40text%3D%BB%F9%BF%C699%40%40tjvalue%3D044&srchStr2=&cxid2=3904&cxczf2=%3D&3904=&cxl3=30&ifftj3=0&ztjIdStr3=&ztjListStr3=&srchStr3=&cxid3=3907&cxczf3=%3D&3907=&cxl4=35&ifftj4=0&ztjIdStr4=&ztjListStr4=&srchStr4=&cxid4=3908&cxczf4=%3D&3908=&cxl5=40&ifftj5=0&ztjIdStr5=&ztjListStr5=&srchStr5=&cxid5=3909&cxczf5=like&3909=&cxl6=80&ifftj6=0&ztjIdStr6=&ztjListStr6=&srchStr6=&cxid6=3927&cxczf6=%3E%3D&3927=&cxl7=85&ifftj7=0&ztjIdStr7=&ztjListStr7=&srchStr7=&cxid7=3928&cxczf7=%3C%3D&3928=&cxl8=120&ifftj8=0&ztjIdStr8=&ztjListStr8=&srchStr8=&cxid8=3916&cxczf8=%3D&3916=&cxl9=130&ifftj9=0&ztjIdStr9=&ztjListStr9=&srchStr9=&cxid9=3912&cxczf9=%3D&3912=&cxl10=140&ifftj10=0&ztjIdStr10=&ztjListStr10=&srchStr10=&cxid10=3913&cxczf10=like&3913=&cxl11=230&ifftj11=0&ztjIdStr11=&ztjListStr11=&srchStr11=%CC%E1%BD%BB&cxid11=3919&cxczf11=%3D&3919=%CC%E1%BD%BB&cxl12=260&ifftj12=0&ztjIdStr12=&ztjListStr12=&srchStr12=%CA%C7&cxid12=3926&cxczf12=%3D&3926=%CA%C7&pxlm=3904&pxlx=asc&xsl=3902&xsl=3903&xsl=3907&xsl=3908&xsl=3909&xsl=3910&xsl=3917&xsl=3918&xsl=3911&xsl=3916&xsl=3912&xsl=3913&xsl=3914&input_go="
        '''
        hard_code1 = "linkSql=&is_input_orderItem_0=3902%5Easc&is_input_orderItem_1=3929%5Easc&is_input_orderItem_2=3908%5Easc&is_input_orderItem_3=3918%5Easc&is_input_orderItemCount=4&id=243&currentPage="
        hard_code2 = "&yhxzl=&pageCount=13&hid_div_disp=none&tjsize=13&cxl0=10&ifftj0=0&ztjIdStr0=&ztjListStr0=&srchStr0=&cxid0=3902&cxczf0=%3D&3902=&cxl1=20&ifftj1=0&ztjIdStr1=&ztjListStr1=&srchStr1=&cxid1=3903&cxczf1=like&3903=&cxl2=25&ifftj2=1&ztjIdStr2=3907&srchStr2=&cxid2=3904&cxczf2=%3D&3904=&cxl3=30&ifftj3=0&ztjIdStr3=&ztjListStr3=&srchStr3=&cxid3=3907&cxczf3=%3D&3907=&cxl4=35&ifftj4=0&ztjIdStr4=&ztjListStr4=&ztjIdStr8=&ztjListStr8=&srchStr8=&cxid8=3916&cxczf8=%3D&3916=&cxl9=130&ifftj9=0&ztjIdStr9=&ztjListStr9=&srchStr9=&cxid9=3912&cxczf9=%3D&3912=&cxl10=140&ifftj10=0&ztjIdStr10=&ztjListStr10=&srchStr10=&cxid10=3913&cxczf10=like&3913=&cxl11=230&ifftj11=0&ztjIdStr11=&ztjListStr11=&srchStr11=%CC%E1%BD%BB&cxid11=3919&cxczf11=%3D&3919=%CC%E1%BD%BB&cxl12=260&ifftj12=0&ztjIdStr12=&ztjListStr12=&srchStr12=%CA%C7&cxid12=3926&cxczf12=%3D&3926=%CA%C7&pxlm=3904&pxlx=asc&xsl=3902&xsl=3903&xsl=3907&xsl=3908&xsl=3909&xsl=3910&xsl=3917&xsl=3918&xsl=3911&xsl=3916&xsl=3912&xsl=3913&xsl=3914&input_go="
        '''

        cj = cookielib.CookieJar()
        httpHandler = urllib2.HTTPHandler(debuglevel=1)
#opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cj), httpHandler)
        opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cj))
        opener.addheaders = [
            ("User-agent", "User-Agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.106 Safari/537.36")
            ]

        op = opener.open("http://info.tsinghua.edu.cn")
        data = urllib.urlencode({
            "userName": username,
            "password": password,
            "x": "44",
            "y": "11"})
        op = opener.open("https://info.tsinghua.edu.cn:443/Login", data)
        op = opener.open("http://info.tsinghua.edu.cn/tag.6cf93c0688bc4571.render.userLayoutRootNode.uP?uP_sparam=focusedTabID&focusedTabID=8&uP_sparam=mode&mode=view")
        text = op.read()

        url_re = re.compile('<iframe.*src="(.*)".*id=.*>')
        url_ex = url_re.search(text)
        url = url_ex.group(1)


        for i in range(len(url)-4):
            if url[i:i+4] == 'amp;':
                url = url[0:i]+url[i+4:]
                break
        op = opener.open(url)



        def number(x):
            if len(x) != 10:
                return False
            for i in range(len(x)):
                if x[i] < '0' or x[i] > '9':
                    return False
            return True


        dic = {}
        #for x in range(1, 2):
        for x in range(1, 14):
            print x
            data = hard_code1 + str(x) + hard_code2 + str(x)
            op = opener.open("http://jxxxfw.cic.tsinghua.edu.cn/search.do2", data)
            data = op.read().decode('gbk') # data is unicode here
            d = pq(data)
            a = [pq(item).text().encode("utf-8") for item in d("body")("td")]

            for i in range(len(a)):
                if number(a[i]):
                    st = i
                    break

            while number(a[st]):
                p_info = (a[st], a[st+1], a[st+2])
                if not dic.has_key(p_info):
                    dic[p_info] = []
                dic[p_info].append((a[st+3], a[st+4], a[st+5], a[st+6], a[st+7], a[st+8], a[st+9], a[st+10], a[st+11], a[st+12]))
                st += 13

        get_db().execute('''DROP TABLE GRADES;''')
        get_db().execute('''CREATE TABLE GRADES
                (ID TEXT PRIMARY KEY,
                SID TEXT,
                NAME TEXT,
                CLASS TEXT,
                CID TEXT,
                CNAME TEXT,
                CTAG TEXT,
                GRADE TEXT,
                TIME TEXT,
                CREDIT TEXT,
                CATEGORY TEXT,
                PLACE TEXT,
                PROF TEXT,
                TAG TEXT);''')
        total = 0
        for x in dic:
            for y in dic[x]:
                get_db().execute("INSERT INTO GRADES VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);", (unicode(total), x[0].decode('utf-8'), x[1].decode('utf-8'), x[2].decode('utf-8'), y[0].decode('utf-8'), y[1].decode('utf-8'), y[2].decode('utf-8'), y[3].decode('utf-8'), y[4].decode('utf-8'), y[5].decode('utf-8'), y[6].decode('utf-8'), y[7].decode('utf-8'), y[8].decode('utf-8'), y[9].decode('utf-8')))
                total = total + 1

        cursor = get_db().execute("SELECT * from GRADES")
        data = [make_dict(cursor, row) for row in cursor.fetchall()]
        flag = [True] * len(data)
        for j in range(len(data)):
            if data[j]['TAG'] == u'重修':
                for i in range(j):
                    if data[i]['SID'] == data[j]['SID'] and data[i]['CID'] == data[j]['CID']:
                        flag[i] = False
        p_data = []
        for i in range(len(flag)):
            if flag[i]:
                p_data.append(data[i])
        dt = {}
        pt = {} 
        py = {}
        for item in p_data:
            SID = item['SID']
            CLASS = item['CLASS']
            idx = (SID, CLASS)
            if not dt.has_key(idx):
                dt[idx] = [0.0, 0, 0.0, 0, 0.0]
                pt[idx] = {}
                py[idx] = {}
            if item['GRADE'] == u'A':
                item['GRADE'] = '4.0'
            elif item['GRADE'] == u'A-':
                item['GRADE'] = '3.7'
            elif item['GRADE'] == u'B+':
                item['GRADE'] = '3.3'
            elif item['GRADE'] == u'B':
                item['GRADE'] = '3.0'
            elif item['GRADE'] == u'B-':
                item['GRADE'] = '2.7'
            elif item['GRADE'] == u'C+':
                item['GRADE'] = '2.3'
            elif item['GRADE'] == u'C':
                item['GRADE'] = '2.0'
            elif item['GRADE'] == u'C-':
                item['GRADE'] = '1.5'
            elif item['GRADE'] == u'D+':
                item['GRADE'] = '1.0'
            elif item['GRADE'] == u'D-':
                item['GRADE'] = '1.0'
            elif item['GRADE'] == u'F':
                item['GRADE'] = '0.0'
            elif not item['GRADE'].isdigit():
                continue
            year = int(item['TIME'][:4])
            date = item['TIME'][4:]
            season = None
            if date >= u'0501' and date <= u'0715':
                season = False 
            elif date >= u'1201':
                season = True
            elif date <= u'0301':
                season = True
                year -= 1

            credit = int(item['CREDIT'])
            total = float(item['GRADE']) * credit
            ''' True for autumn and False for Spring ''' 
            if season is not None:
                if not pt[idx].has_key((year, season)):
                    pt[idx][(year, season)] = [0.0, 0]
                pt[idx][(year, season)][0] += total
                pt[idx][(year, season)][1] += credit
                if not season:
                    year -= 1
                if not py[idx].has_key(year):
                    py[idx][year] = [0.0, 0]
                py[idx][year][0] += total
                py[idx][year][1] += credit

            if item['CATEGORY'] == u'必修' or item['CATEGORY'] == u'限选':
                dt[idx][0] += total
                dt[idx][1] += credit
            dt[idx][2] += total
            dt[idx][3] += credit
            dt[idx][4] += total
        for x in pt:
            for y in pt[x]:
                pt[x][y] = pt[x][y][0] / pt[x][y][1]
        for x in py:
            for y in py[x]:
                py[x][y] = py[x][y][0] / py[x][y][1]
        get_db().execute('''DROP TABLE GPA''')
        get_db().execute('''CREATE TABLE GPA
                (SID TEXT PRIMARY KEY,
                CLASS TEXT,
                SEASON_DICT TEXT,
                YEAR_DICT TEXT,
                REQ_GPA FLOAT,
                ALL_GPA FLOAT,
                TOT_GPA FLOAT);''')
        for idx in dt:
            if dt[idx][1] == 0:
                dt[idx][1] = 1
            if dt[idx][3] == 0:
                dt[idx][3] = 1
            get_db().execute("INSERT INTO GPA VALUES (?, ?, ?, ?, ?, ?, ?);", [idx[0], idx[1], pickle.dumps(pt[idx]), pickle.dumps(py[idx]), float(dt[idx][0])/dt[idx][1], float(dt[idx][2])/dt[idx][3], dt[idx][4]])
        
        print 'before commit '
        get_db().commit()
        return True
    except:
        return False

def make_excel():
    cursor = get_db().execute("SELECT * from GRADES;")
    grades  = [make_dict(cursor, row) for row in cursor.fetchall()]
    cursor = get_db().execute("SELECT * from GPA;")
    gpas = [make_dict(cursor, row) for row in cursor.fetchall()]
    cursor = get_db().execute("SELECT CHOICE from USERS;")
    users = [make_dict(cursor, row) for row in cursor.fetchall()]

    grade_set = set()
    for idx in gpas:
        grade_set.add(idx['CLASS'][:-1])
    grade_list = sorted(list(grade_set))

    w = xlwt.Workbook(encoding='utf-8')
    ws = []
    grade_to_idx = {}
    for item in grade_list:
        grade_to_idx[item] = len(ws)
        ws.append(w.add_sheet(item + u"所有成绩"))
    row = [1] * len(grade_list)

    for i in range(len(ws)):
        ws[i].write(0, 0, "学号")
        ws[i].write(0, 1, "姓名")
        ws[i].write(0, 2, "教学班级")
        ws[i].write(0, 3, "课程号")
        ws[i].write(0, 4, "课程名")
        ws[i].write(0, 5, "课序号")
        ws[i].write(0, 6, "成绩")
        ws[i].write(0, 7, "考试时间")
        ws[i].write(0, 8, "学分")
        ws[i].write(0, 9, "课程属性")
        ws[i].write(0, 10, "教室号")
        ws[i].write(0, 11, "教师名")
        ws[i].write(0, 12, "重修补考标志")

    for item in grades:
        idx = grade_to_idx[item['CLASS'][:-1]]
        ws[idx].write(row[idx], 0, item['SID'])
        ws[idx].write(row[idx], 1, item['NAME'])
        ws[idx].write(row[idx], 2, item['CLASS'])
        ws[idx].write(row[idx], 3, item['CID'])
        ws[idx].write(row[idx], 4, item['CNAME'])
        ws[idx].write(row[idx], 5, item['CTAG'])
        ws[idx].write(row[idx], 6, item['GRADE'])
        ws[idx].write(row[idx], 7, item['TIME'])
        ws[idx].write(row[idx], 8, item['CREDIT'])
        ws[idx].write(row[idx], 9, item['CATEGORY'])
        ws[idx].write(row[idx], 10, item['PLACE'])
        ws[idx].write(row[idx], 11, item['PROF'])
        ws[idx].write(row[idx], 12, item['TAG'])
        row[idx] += 1


    for cnt_grade in grade_list:
        ws = w.add_sheet(cnt_grade + u'GPA')
        ws.write(0, 0, "学号")
        ws.write(0, 1, "姓名")
        ws.write(0, 2, "教学班级")
        ws.write(0, 3, "必修限选GPA")
        ws.write(0, 4, "必修限选GPA排名")
        ws.write(0, 5, "整体GPA")
        ws.write(0, 6, "整体GPA排名")
        ws.write(0, 7, "整体学分和")
        ws.write(0, 8, "整体学分和排名")
        a, b, c = [], [], []
        for gpa in gpas:
            if gpa['CLASS'][:-1] == cnt_grade:
                a.append(gpa['REQ_GPA'])
                b.append(gpa['ALL_GPA'])
                c.append(gpa['TOT_GPA'])
        a = sorted(a, reverse=True)
        b = sorted(b, reverse=True)
        c = sorted(c, reverse=True)
        row = 1
        for gpa in gpas:
            if gpa['CLASS'][:-1] == cnt_grade:
                ws.write(row, 0, gpa['SID'])
                ws.write(row, 1, get_db().execute("SELECT NAME from GRADES where SID = ?;", [gpa['SID']]).fetchone()[0])
                ws.write(row, 2, gpa['CLASS'])
                ws.write(row, 3, gpa['REQ_GPA'])
                ws.write(row, 4, a.index(gpa['REQ_GPA'])+1)
                ws.write(row, 5, gpa['ALL_GPA'])
                ws.write(row, 6, b.index(gpa['ALL_GPA'])+1)
                ws.write(row, 7, gpa['TOT_GPA'])
                ws.write(row, 8, c.index(gpa['TOT_GPA'])+1)
                row += 1
        l_tags = 9
        season_list = []
        year_list = []
        for gpa in gpas:
            if gpa['CLASS'][:-1] == cnt_grade:
                season_dict = pickle.loads(gpa['SEASON_DICT'])
                year_dict = pickle.loads(gpa['YEAR_DICT'])
                for idx in season_dict:
                    if not idx in season_list:
                        season_list.append(idx)
                for idx in year_dict:
                    if not idx in year_list:
                        year_list.append(idx)

        season_list = sorted(season_list)
        year_list = sorted(year_list)

        ''' ds[class][season] = list '''
        ''' dy[class][year]   = list '''
        ds, dy = {}, {}
        for gpa in gpas:
            cnt_class = gpa['CLASS']
            if cnt_class[:-1] == cnt_grade:
                if not ds.has_key(cnt_class):
                    ds[cnt_class], dy[cnt_class] = {}, {}
            ds['all'], dy['all'] = {}, {}
        for idx in ds:
            for season in season_list:
                ds[idx][season] = []
        for idx in dy:
            for year in year_list:
                dy[idx][year] = []
        
        talent = cnt_grade[0] == u'基'
        for gpa in gpas:
            cnt_class = gpa['CLASS']
            if cnt_class[:-1] == cnt_grade:
                season_dict = pickle.loads(gpa['SEASON_DICT'])
                year_dict = pickle.loads(gpa['YEAR_DICT'])
                for season in season_dict:
                    ds[cnt_class][season].append(season_dict[season])
                    ds['all'][season].append(season_dict[season])
                for year in year_dict:
                    dy[cnt_class][year].append(year_dict[year])
                    dy['all'][year].append(year_dict[year])

        for to_sort in [ds, dy]:
            for idx1 in to_sort:
                for idx2 in to_sort[idx1]:
                    to_sort[idx1][idx2] = sorted(to_sort[idx1][idx2], reverse=True)

        if talent:
            for season in season_list:
                ws.write(0, l_tags, "%d年%s学期总GPA" % (season[0], season and "秋季" or "春季"))
                ws.write(0, l_tags+1, "%d年%s学期总GPA基科班排名" % (season[0], season and "秋季" or "春季"))
                row = 1
                for gpa in gpas:
                    cnt_class = gpa['CLASS']
                    if cnt_class[:-1] == cnt_grade:
                        season_dict = pickle.loads(gpa['SEASON_DICT'])
                        if season_dict.has_key(season):
                            ws.write(row, l_tags, season_dict[season])
                            ws.write(row, l_tags+1, ds[cnt_class][season].index(season_dict[season])+1)
                        row += 1
                l_tags += 2
            for year in year_list:
                ws.write(0, l_tags, "%d-%d学年总GPA" % (year, year+1))
                ws.write(0, l_tags+1, "%d-%d学年总GPA班内排名" % (year, year+1))
                row = 1
                for gpa in gpas:
                    cnt_class = gpa['CLASS']
                    if cnt_class[:-1] == cnt_grade:
                        year_dict = pickle.loads(gpa['YEAR_DICT'])
                        if year_dict.has_key(year):
                            ws.write(row, l_tags, year_dict[year])
                            ws.write(row, l_tags+1, dy[cnt_class][year].index(year_dict[year])+1)
                        row += 1
                l_tags += 2
        else:
            for season in season_list:
                ws.write(0, l_tags, "%d年%s学期总GPA" % (season[0], season and "秋季" or "春季"))
                ws.write(0, l_tags+1, "%d年%s学期总GPA班内排名" % (season[0], season and "秋季" or "春季"))
                ws.write(0, l_tags+2, "%d年%s学期总GPA年纪排名" % (season[0], season and "秋季" or "春季"))
                row = 1
                for gpa in gpas:
                    cnt_class = gpa['CLASS']
                    if cnt_class[:-1] == cnt_grade:
                        season_dict = pickle.loads(gpa['SEASON_DICT'])
                        if season_dict.has_key(season):
                            ws.write(row, l_tags, season_dict[season])
                            ws.write(row, l_tags+1, ds[cnt_class][season].index(season_dict[season])+1)
                            ws.write(row, l_tags+2, ds['all'][season].index(season_dict[season])+1)
                        row += 1
                l_tags += 3
            for year in year_list:
                ws.write(0, l_tags, "%d-%d学年总GPA" % (year, year+1))
                ws.write(0, l_tags+1, "%d-%d学年总GPA班内排名" % (year, year+1))
                ws.write(0, l_tags+2, "%d-%d学年总GPA年级排名" % (year, year+1))
                row = 1
                for gpa in gpas:
                    cnt_class = gpa['CLASS']
                    if cnt_class[:-1] == cnt_grade:
                        year_dict = pickle.loads(gpa['YEAR_DICT'])
                        if year_dict.has_key(year):
                            ws.write(row, l_tags, year_dict[year])
                            ws.write(row, l_tags+1, dy[cnt_class][year].index(year_dict[year])+1)
                            ws.write(row, l_tags+2, dy['all'][year].index(year_dict[year])+1)
                        row += 1
                l_tags += 3
        
    w.save('static/heiheihei.xls')


class User(UserMixin):
    def __init__(self, username, password, choice, *args, **kwargs):
        self.id = username
        self.password = password
        self.choice = choice
        UserMixin.__init__(self, *args, **kwargs)


class LoginForm(Form):
    username = TextField('Username', [validators.Required()])
    password = PasswordField('Password', [validators.Required()])

    def __init__(self, *args, **kwargs):
        kwargs['csrf_enabled'] = False
        Form.__init__(self, *args, **kwargs)
        self.user = None

    def validate(self):
        if not Form.validate(self):
            return False

        if self.username.data == 'admin' and self.password.data == 'qwer1234':
            return True
        cursor = get_db().execute("SELECT * from USERS where USERNAME = ?;", [self.username.data])
        userrow = make_dict(cursor, cursor.fetchone())
        if userrow is None:
            print self.username.errors
            #self.username.errors.append('Unknown username')
            return False

        if get_db().execute("SELECT CLASS from GPA where SID = ?;", [userrow['USERNAME']]).fetchone()[0][2] == u'2':
            ''' 二字班的同学说不要看到成绩 '''
            return False

        if self.password.data != userrow['PASSWORD']:
            #self.password.errors.append('Invalid password')
            return False

        self.user = User(username=userrow['USERNAME'], password=userrow['PASSWORD'], choice=userrow['CHOICE'])
        return True

class ChangeForm(Form):
    username = TextField('Username', [validators.Required()])
    password = PasswordField('Password', [validators.Required()])
    chanword = PasswordField('Chanword', [validators.Required()])
    confword = PasswordField('Confword', [validators.Required()])

    def __init__(self, *args, **kwargs):
        kwargs['csrf_enabled'] = False
        Form.__init__(self, *args, **kwargs)

class ChoiceForm(Form):
    choice_list1 = SelectField(u'选择志愿')
    choice_list2 = SelectField(u'选择志愿')
    choice_list3 = SelectField(u'选择志愿')
    submit = SubmitField('Submit')

    def __init__(self, *args, **kwargs):
        Form.__init__(self, *args, **kwargs)
        self.choice_list1.choices = self.choice_list2.choices = self.choice_list3.choices = [
                (u'', u'---------'),
                (u'本系推研', u'本系推研'),
                (u'本系直博', u'本系直博'),
                (u'出国', u'出国'),
                (u'深研院', u'深研院'),
                (u'外系推研', u'外系推研'),
                (u'外校推研', u'外校推研'),
                (u'其它', u'其它')]

@login_manager.user_loader
def load_user(userid):
    cursor = get_db().execute("SELECT * from USERS where USERNAME = ?;", [userid])
    userrow = cursor.fetchone()
    return User(username=userrow[0], password=userrow[1], choice=userrow[2])

@app.route('/index', methods=['GET', 'POST'])
@login_required
def index():
    if request.method == 'POST':
        up1 = request.form['choice_list1']
        up2 = request.form['choice_list2']
        up3 = request.form['choice_list3']
        get_db().execute("UPDATE USERS SET CHOICE = ? where USERNAME = ?;", [up1+u' '+up2+u' '+up3, current_user.id])
        get_db().commit()

    ''' fetch data from GRADES '''
    cursor = get_db().execute("SELECT * from GRADES where SID = ?;", [current_user.id])
    data = [make_dict(cursor, row) for row in cursor.fetchall()]
    cursor = get_db().execute("SELECT * from GPA where SID = ?;", [current_user.id])
    gpa = make_dict(cursor, cursor.fetchone())
    cursor = get_db().execute("SELECT CHOICE from USERS where USERNAME = ?;", [current_user.id])
    choice = make_dict(cursor, cursor.fetchone())

    talent = gpa['CLASS'][0] == u'基'

    ''' ranks '''
    cursor = get_db().execute("SELECT * from GPA;")
    raw_data = [make_dict(cursor, row) for row in cursor.fetchall()]

    ''' season '''
    season_dict = pickle.loads(gpa['SEASON_DICT'])
    year_dict = pickle.loads(gpa['YEAR_DICT'])
    dt = {}
    dt1 = {}
    dt2 = {}
    dy = {}
    dy1 = {}
    dy2 = {}
    for item in season_dict:
        dt[item] = []
        dt1[item] = []
        dt2[item] = []
    for item in year_dict:
        dy[item] = []
        dy1[item] = []
        dy2[item] = []

    for item in raw_data:
        if item['CLASS'][2] == gpa['CLASS'][2] and item['CLASS'][0] == gpa['CLASS'][0]:
            ''' same grade and same talent '''
            cnt_season_dict = pickle.loads(item['SEASON_DICT'])
            cnt_year_dict = pickle.loads(item['YEAR_DICT'])
            for idx in season_dict:
                if not cnt_season_dict.has_key(idx):
                    ''' although this is impossible '''
                    continue
                if talent:
                    dt[idx].append(cnt_season_dict[idx])
                else:
                    if item['CLASS'] == gpa['CLASS']:
                        dt1[idx].append(cnt_season_dict[idx])
                    dt2[idx].append(cnt_season_dict[idx])
            for idx in year_dict:
                if not cnt_year_dict.has_key(idx):
                    ''' although this is impossible '''
                    continue
                if talent:
                    dy[idx].append(cnt_year_dict[idx])
                else:
                    if item['CLASS'] == gpa['CLASS']:
                        dy1[idx].append(cnt_year_dict[idx])
                    dy2[idx].append(cnt_year_dict[idx])

    rt = {}
    rt1 = {}
    rt2 = {}
    ry = {}
    ry1 = {}
    ry2 = {}
    if talent:
        for x in dt:
            rt[x] = sorted(dt[x])[::-1].index(season_dict[x]) + 1
        for x in dy:
            ry[x] = sorted(dy[x])[::-1].index(year_dict[x]) + 1
    else:
        for x in dt1:
            rt1[x] = sorted(dt1[x])[::-1].index(season_dict[x]) + 1
            rt2[x] = sorted(dt2[x])[::-1].index(season_dict[x]) + 1
        for x in dy1:
            ry1[x] = sorted(dy1[x])[::-1].index(year_dict[x]) + 1
            ry2[x] = sorted(dy2[x])[::-1].index(year_dict[x]) + 1

    talent_req = []
    talent_all = []
    talent_tot = []
    normal_req = []
    normal_all = []
    normal_tot = []
    for item in raw_data:
        if item['CLASS'][2] == gpa['CLASS'][2]:
            if item['CLASS'][0] == u'基':
                talent_req.append(item['REQ_GPA'])
                talent_all.append(item['ALL_GPA'])
                talent_tot.append(item['TOT_GPA'])
            else:
                normal_req.append(item['REQ_GPA'])
                normal_all.append(item['ALL_GPA'])
                normal_tot.append(item['TOT_GPA'])
    talent_req = sorted(talent_req)[::-1]
    talent_all = sorted(talent_all)[::-1]
    talent_tot = sorted(talent_tot)[::-1]
    normal_req = sorted(normal_req)[::-1]
    normal_all = sorted(normal_all)[::-1]
    normal_tot = sorted(normal_tot)[::-1]
    t6_req = talent_req[int(len(talent_req)*0.6)-1]
    t4_req = talent_req[int(len(talent_req)*0.4)-1]
    t8_req = talent_req[int(len(talent_req)*0.8)-1]
    t4_all = talent_all[int(len(talent_all)*0.4)-1]
    t8_all = talent_all[int(len(talent_all)*0.8)-1]
    t4_tot = talent_tot[int(len(talent_tot)*0.4)-1]
    t8_tot = talent_tot[int(len(talent_tot)*0.8)-1]
    n4_req = normal_req[int(len(normal_req)*0.4)-1]
    n8_req = normal_req[int(len(normal_req)*0.8)-1]
    n4_all = normal_all[int(len(normal_all)*0.4)-1]
    n8_all = normal_all[int(len(normal_all)*0.8)-1]
    n4_tot = normal_tot[int(len(normal_tot)*0.4)-1]
    n8_tot = normal_tot[int(len(normal_tot)*0.8)-1]

    form = ChoiceForm()
    choices = choice['CHOICE'].split(' ')
    if len(choices) == 3:
        form.choice_list1.data, form.choice_list2.data, form.choice_list3.data = choices
    else:
        form.choice_list1.data = form.choice_list2.data = form.choice_list3.data = u''

    return render_template('index.html',
            form=form,
            data=data,
            gpa=gpa,
            talent=talent,
            t6_req=t6_req,
            t4_req=t4_req,
            t8_req=t8_req,
            t4_all=t4_all,
            t8_all=t8_all,
            t4_tot=t4_tot,
            t8_tot=t8_tot,
            n4_req=n4_req,
            n8_req=n8_req,
            n4_all=n4_all,
            n8_all=n8_all,
            n4_tot=n4_tot,
            n8_tot=n8_tot,
            rt=rt,
            rt1=rt1,
            rt2=rt2,
            ry=ry,
            ry1=ry1,
            ry2=ry2,
            season_tuple=sorted(season_dict.iteritems(), key=lambda d:d[0]),
            year_tuple=sorted(year_dict.iteritems(), key=lambda d:d[0]),
            )

@app.route('/login', methods=['GET', 'POST'])
def login():
    form = LoginForm()
    if form.validate_on_submit():
        if form.user == None:
            return redirect(url_for('super'))
        login_user(form.user)
        return redirect(url_for('index'))

    return render_template('login.html', form=form)

@app.route('/update', methods=['GET', 'POST'])
def update():
    if init_the_database():
        return redirect(url_for('super'))
    else:
        return 'update failed'
    

@app.route('/super', methods=['GET', 'POST'])
def super():
    make_excel()
    cursor = get_db().execute("SELECT * from USERS;", [])
    choices = [make_dict(cursor, row) for row in cursor.fetchall()]
    for item in choices:
        ''' USERNAME CHOICE '''
        item['NAME'] = get_db().execute("SELECT NAME from GRADES where SID = ?;", [item['USERNAME']]).fetchone()[0]

        ret = item['CHOICE'].split(' ')
        if len(ret) == 3:
            item['CHOICE1'] = ret[0]
            item['CHOICE2'] = ret[1]
            item['CHOICE3'] = ret[2]
        else:
            item['CHOICE1'] = item['CHOICE1'] = item['CHOICE1'] = u''
    return render_template('super.html', choices=choices)

@app.route('/', methods=['GET', 'POST'])
def main():
    return redirect(url_for('login'))

@app.route('/heiheihei.xls', methods = ['POST'])
def download():
    if request.method == 'POST':
        return send_from_directory('static', 'heiheihei.xls', as_attachment=True)
    else:
        return 'Download failed'

@app.route('/change', methods=['GET', 'POST'])
def change():
    form = ChangeForm()
    if form.validate_on_submit():
        cursor = get_db().execute("SELECT * from USERS where USERNAME = ?;", [form.username.data])
        userrow = cursor.fetchone()
        if userrow is None:
            return render_template('change.html', form=form)
        if form.password.data != userrow[1]:
            return render_template('change.html', form=form)
        if form.chanword.data != form.confword.data:
            return render_template('change.html', form=form)
        get_db().execute("UPDATE USERS SET PASSWORD = ? where USERNAME = ?;", [form.chanword.data, form.username.data])
        get_db().commit()
        return render_template('success.html')

    return render_template('change.html', form=form)

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return 'logout ok'

if __name__ == '__main__':
    app.run(host='0.0.0.0')
    #app.run()
