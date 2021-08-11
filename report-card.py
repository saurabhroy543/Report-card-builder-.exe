import xlrd
from reportlab.graphics.charts.legends import (Legend)
from reportlab.lib.validators import Auto
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Table, SimpleDocTemplate, TableStyle, Image
import datetime

from reportlab.graphics.shapes import Drawing, String

from reportlab.graphics.charts.piecharts import (Pie)


def gePieChart(data, name):
    chart = Pie()
    chart.data = data
    chart.x = 50
    chart.y = 5
    chart.labels = ['Correct', 'Incorrect', 'Unattempted']
    chart.slices[0].fillColor = colors.green
    chart.slices[0].popout = 3
    title = String(
        5, 120, name,
        fontSize=20
    )
    legend = Legend()
    legend.x = 180
    legend.y = 80
    legend.alignment = 'right'

    legend.colorNamePairs = Auto(obj=chart)

    drawing = Drawing(240, 120)
    drawing.add(title)
    drawing.add(chart)
    drawing.add(legend)

    return drawing


def piechartdata(l):
    un = 0
    c = 0
    inn = 0
    for i in l:
        if i == 'Unattempted':
            un = un + 1
        elif i == 'Correct':
            c = c + 1
        elif i == 'Incorrect':
            inn = inn + 1
    co = [c, inn, un]
    return co


path = "resource/Dummy.xlsx"

workbook = xlrd.open_workbook(path)  # opening the excel file
worksheet = workbook.sheet_by_index(0)  # selecting the sheet


def createpdf(name, lis, lis2, lis3, lis4, lis0, new, new1):
    file = '.pdf'
    filename = name + file
    pdf = SimpleDocTemplate(filename, pagesize=A4)
    table2 = Table(lis)
    table1 = Table(lis2)
    table4 = Table(lis4)
    table3 = Table(lis3)
    table0 = Table(lis0)
    style3 = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('TEXTCOLOR', (0, 0), (0, 0), colors.red),
        ('TEXTCOLOR', (1, 0), (1, 0), colors.purple),
        ('TEXTCOLOR', (2, 0), (2, 0), colors.red),
        ('TOPPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 0), (-1, 0), colors.beige),
    ])
    style = TableStyle([
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.blue),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('FONTNAME', (0, 0), (-1, 0), 'Courier-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('BACKGROUND', (0, 0), (-1, 0), colors.beige),
    ])
    style0 = TableStyle([
        ('TEXTCOLOR', (1, 0), (1, 0), colors.purple),
        ('FONTSIZE', (1, 0), (1, 0), 19),
        ('BOTTOMPADDING', (1, 0), (1, 0), 15),
        ('ALIGN', (0, 1), (1, 1), 'CENTER'),
    ])
    table0.setStyle(style0)
    table2.setStyle(style)
    table1.setStyle(style)
    table4.setStyle(style)
    table3.setStyle(style3)
    piechart = gePieChart(new, 'your outcome')
    world = gePieChart(new1, 'world avrage outcome')
    chart = Table([[piechart, world]])

    elems = [table0, table1, table4, table2, table3, chart]
    pdf.build(elems)
    return "done"


# world average attempts:
wa = []
for i in range(2, worksheet.nrows):
    wa.append(worksheet.cell_value(i, 16))

k = ''
n = 0
for i in range(2, worksheet.nrows):
    y = worksheet.cell_value(i, 2)
    if k != y:
        k = y
        n = n + 1
j = 2  # n is the number of student whose data is inserted in the excel file
while n != 0:  # looping until the number of student is zero
    count = 0
    total = 0
    round = int(worksheet.cell_value(j, 1))
    firstname = (worksheet.cell_value(j, 2))
    lastname = worksheet.cell_value(j, 3)
    fullname = worksheet.cell_value(j, 4)
    reg = worksheet.cell_value(j, 5)
    grade = int(worksheet.cell_value(j, 6))
    gender = worksheet.cell_value(j, 8)

    # date formatting for date of birth
    dob = worksheet.cell_value(j, 9)
    c_date = xlrd.xldate_as_tuple(dob, workbook.datemode)
    n_date = datetime.datetime(*c_date).strftime("%d/%m/%y")
    remark = worksheet.cell_value(j, 19)

    # picture to the pdf with school name and logo
    head = []
    h = []
    pic = 'resource/' + fullname + '.png'
    picture = Image(pic, width=40, height=40)
    logo = Image('resource/logo.png', width=30, height=30)
    school = worksheet.cell_value(j, 7)
    h.append(logo)
    h.append(school)
    h.append(picture)
    head.append(h)

    # round and exam name
    r = ['', 'Mid-Term:Round', round]
    head.append(r)

    # creating required list for bio of student
    lsii = [['First Name', 'Last Name', 'Full Name', 'registration No.', 'D.O.B', 'Grade', 'Gender']]
    ts, ls, ts1 = [], [], ['REMARK:']

    ls.append(firstname)
    ls.append(lastname)
    ls.append(fullname)
    ls.append(reg)
    ls.append(n_date)
    ls.append(grade)
    ls.append(gender)
    lsii.append(ls)

    # list for exam related info
    exam = [['City of Residence', 'Date and time of Exam', 'Country of Residence']]
    ex = []
    city = worksheet.cell_value(j, 10)
    exam_date = worksheet.cell_value(j, 11)
    country = worksheet.cell_value(j, 12)
    ex.append(city)
    ex.append(exam_date)
    ex.append(country)
    exam.append(ex)
    correct = []
    # list creation and extraction of report of exam of students
    L = [['Question No.', 'What you marked', 'Correct Answer', 'Outcome', 'Score if correct', 'Your score']]
    for i in range(2, worksheet.nrows):

        y = worksheet.cell_value(i, 2)
        if y == firstname:
            correct.append(worksheet.cell_value(i, 16))
            x = int(worksheet.cell_value(i, 18))
            lv = int(worksheet.cell_value(i, 17))
            g = []
            for o in range(13, 19):
                g.append(worksheet.cell_value(i, o))
            L.append(g)
            total = total + lv
            count = count + x
            j += 1
    # print("Total score of %s is %s out of %s %s" % (firstname, count, total, j))

    ts1.append(remark)
    ts1.append('marks obtained')
    ts1.append(count)
    ts1.append('out of')
    ts1.append(total)
    ts.append(ts1)
    da = piechartdata(correct)
    worldd = piechartdata(wa)
    v = createpdf(firstname, L, lsii, ts, exam, head, da, worldd)
    n -= 1
