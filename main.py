from flask_mail import Mail
import os
from flask import Flask,render_template,request,session, redirect
from flask_sqlalchemy import SQLAlchemy
import xlrd
from PIL import Image, ImageDraw, ImageFont
from werkzeug.utils import secure_filename
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = ".\\static\\files"
app.config.update(
    MAIL_SERVER='smtp.gmail.com',
    MAIL_PORT='465',
    MAIL_USE_SSL=True,
    # Enter Gmail Id and password
    MAIL_USERNAME="Enter Gmail id",
    MAIL_PASSWORD="Enter Gmail Password"
)
# Enter Gmail Id and password
MAIL_USERNAME="Enter Gmail id",
MAIL_PASSWORD="Enter Gmail Password"
mail = Mail(app)

def collect_Data(loc):
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    i = 1
    participants = []
    while i < sheet.nrows:
        name = sheet.row_values(i)[1]
        year = sheet.row_values(i)[2]
        branch = sheet.row_values(i)[3]
        event = sheet.row_values(i)[4]
        position = sheet.row_values(i)[5]
        email = sheet.row_values(i)[6]
        l = [name, year, branch, event, position,email]
        participants.append(l)
        # image.save(f'send{sno}.png')
        i = i + 1
    return participants

def cerid(srno):
    pass

def generate_certificate(name, year, branch, event, position, email):
    position = str(int(position))
    date = "25 Nov 2020"
    fontTitle = ImageFont.truetype('title.ttf', size=70)
    fontName = ImageFont.truetype('title.ttf', size=100)
    image = Image.open('certificate.jpg')
    draw = ImageDraw.Draw(image)
    (x, y) = (1600, 1350)
    color = 'rgb(0, 0, 0)'  # black color
    draw.text((x, y), name, fill=color, font=fontName, spacing=6)
    (x, y) = (600, 1600)
    draw.text((x, y), year, fill=color, font=fontTitle, spacing=6)
    (x, y) = (1250, 1600)
    draw.text((x, y), branch, fill=color, font=fontTitle, spacing=6)
    (x, y) = (900, 1700)
    draw.text((x, y), event, fill=color, font=fontTitle, spacing=6)
    (x, y) = (2900, 1600)
    draw.text((x, y), position, fill=color, font=fontTitle, spacing=6)
    (x, y) = (1280, 1790)
    draw.text((x, y), date, fill=color, font=fontTitle, spacing=6)
    image.save(f'{name}.png')
    body = "Dear " + name + "," + "\n" + " Please collect the certificate of " + event + "\n" + \
           "\n" + "Please Check in Attachments"
    msg = MIMEMultipart()
    # Enter Senders Gmail Id
    msg['From'] = "Enter Gmail id"
    msg['To'] = email
    msg['Subject'] = 'Congratulations you got it ... '
    msg.attach(MIMEText(body, 'plain'))
    filename = f'{name}.png'
    attachment = open(filename, 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= " + filename)
    msg.attach(part)
    text = msg.as_string()
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(MAIL_USERNAME, MAIL_PASSWORD)
    server.sendmail(MAIL_USERNAME, email, text)
    server.quit()


excel_filename = ""


@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        global excel_filename
        f = request.files['upl']
        print(f.filename)
        f.save(os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(f.filename)))
        excel_filename = f.filename
        return redirect("/waiting")
    return render_template("index.html")


@app.route("/waiting")
def waiting():
    global excel_filename
    loc = app.config['UPLOAD_FOLDER'] + "\\" + excel_filename
    data = collect_Data(loc)
    for row in data:
        generate_certificate(row[0], row[1], row[2], row[3], row[4], row[5])
    return render_template("waiting.html")


if __name__ == "__main__":
    app.run(debug=True)
