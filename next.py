from flask import Flask, render_template, redirect, url_for, request, make_response
from flask_sqlalchemy import SQLAlchemy
import flask_excel as excel
import os

from pyexcel_webio import make_response_from_tables

basedir = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)
excel.init_excel(app)
app.secret_key = 'some secret salt'
app.config['SQLALCHEMY_DATABASE_URI'] = \
    'sqlite:///' + os.path.join(basedir, 'data1.sqlite?check_same_thread=False')
app.config['SQLALCHEMY_COMMIT_ON_TEARDOWN'] = False
db = SQLAlchemy(app)


def ppw():
    m = HH.query.first()
    n = str(m)
    db.session.commit()
    return n


def pp(id):
    m = HH.query.get(id)
    n = str(m)
    db.session.commit()
    return n


def pw():
    import random
    pas1 = list('1234567890')
    pas2 = list('ABCDEFGHIGKLMNOPQRSTUVYXWZ')
    pas3 = list('abcdefghigklmnopqrstuvyxwz')
    random.shuffle(pas1)
    random.shuffle(pas2)
    random.shuffle(pas3)
    passwd = ''.join(
        [random.choice(pas1) for x in range(4)] + [random.choice(pas2) for x in range(1)] + [random.choice(pas3) for x
                                                                                             in range(1)])
    return passwd


class HH(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    txr = db.Column(db.String(32), nullable=False)

    def __init__(self, txr):
        self.txr = txr

    def __repr__(self):
        return '{}'.format(self.txr)


class Pass(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    pg = db.Column(db.String(32), nullable=False)

    us_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    users = db.relationship('Users', cascade='all,delete', backref=db.backref('tags', lazy=True))

    def __repr__(self):
        return '{}'.format(self.pg)


class Users(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False)
    last_name = db.Column(db.String(255), nullable=False)
    sp = db.Column(db.String(255), nullable=False)
    job = db.Column(db.String(255), nullable=False)
    email = db.Column(db.String(255), nullable=False)
    tel = db.Column(db.String(255), nullable=False)

    def __init__(self, name, last_name, sp, job, email, tel, tags):
        self.name = name.strip()
        self.last_name = last_name.strip()
        self.sp = sp.strip()
        self.job = job.strip()
        self.email = email.strip()
        self.tel = tel.strip()
        self.tags = [
            Pass(pg=tags.strip())
        ]

    def __repr__(self):
        return '<Name {}>'.format(self.name)


db.create_all()


@app.route('/main', methods=['GET', 'POST'])
def main_p():
    return render_template('main2.html', items=Users.query.all())


@app.route('/', methods=['GET', 'POST'])
def main():
    return render_template('main.html')


@app.route('/add', methods=['GET', 'POST'])
def add():
    a = [pw() for x in range(10)]

    return render_template('index.html', a=a)


@app.route('/add_message', methods=['POST'])
def add_message():
    name = request.form['name']
    last_name = request.form['last_name']
    sp = request.form['sp']
    job = request.form['job']
    email = request.form['email']
    tel = request.form['tel']
    tag = ppw()

    db.session.add(Users(name, last_name, sp, job, email, tel, tag))
    db.session.commit()

    return redirect(url_for('main_p'))


@app.route("/export", methods=['GET', 'POST'])
def do_export():
    return make_response_from_tables(db.session, [Users], "xls")


@app.route('/adi/<int:id>', methods=['GET', 'POST'])
def adi(id):
    t = Pass.query.get(id)
    t.pg = pp(id)
    db.session.commit()
    return redirect(url_for('main_p'))


@app.route('/delete/<int:id>')
def delete(id):
    task_to_delete = Pass.query.get_or_404(id)
    try:
        db.session.delete(task_to_delete)
        db.session.commit()
        return redirect(url_for('main_p'))
    except:
        return 'There was a problem deleting that task'


@app.route('/update/<int:id>', methods=['GET', 'POST'])
def update(id):
    task = Users.query.get(id)
    t = Pass.query.get_or_404(id)
    if request.method == 'POST':
        task.name = request.form['name']
        task.last_name = request.form['last_name']
        task.sp = request.form['sp']
        task.job = request.form['job']
        task.email = request.form['email']
        task.tel = request.form['tel']
        t.pg = request.form['tag']
        try:
            db.session.commit()
            return redirect(url_for('main_p'))
        except:
            return 'w'
    else:
        return render_template('update.html', task=task, t=t)


@app.route("/import", methods=['GET', 'POST'])
def doimport():
    if request.method == 'POST':

        def category_init_func(row):
            c = HH(row['txr'])
            c.id = row['id']
            return c

        request.save_book_to_database(field_name='file', session=db.session, tables=[HH], initializers=[category_init_func])
        return redirect(url_for('.handson_table'), code=302)
    return '''
    <!doctype html>
    <title>Upload an excel file</title>
    <h1>Excel file upload (xls, xlsx, ods please)</h1>
    <form action="" method=post enctype=multipart/form-data><p>
    <input type=file name=file><input type=submit value=Upload>
    </form>
    '''


@app.route("/handson_view", methods=['GET'])
def handson_table():
    return excel.make_response_from_tables(db.session, [HH], 'handsontable.html')


if __name__ == "__main__":
    app.run(host='127.0.0.1')
