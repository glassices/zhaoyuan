# -*- coding: utf-8 -*-

from flask import Flask
from flask import g
from flask import url_for, redirect, render_template, request
from flask.ext.wtf import Form
from wtforms import TextField, PasswordField, validators, SelectField, SubmitField
from flask.ext.login import LoginManager, UserMixin, login_user, login_required, current_user
from flask_bootstrap import Bootstrap
import sqlite3
import pickle

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

@app.route('/super', methods=['GET', 'POST'])
def super():
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
