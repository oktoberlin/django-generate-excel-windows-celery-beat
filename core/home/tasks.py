# task.py, this file can be present in any of the apps.
from celery.schedules import crontab

from celery import shared_task
import pandas as pd
import xlsxwriter
import pymysql
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from home.models import generate_excel

def send_mail_task(client, email, start_date, end_date):
    # connecting server database to python
    db = pymysql.connect(user="admin21",password="adminsby21",host="149.129.243.10",port=3307,database="mydepo")
    sql1 = """Select EIRIN.ContNo,ContainerDetails.size,ContainerDetails.type,contCondition,
    Payload,If(Net=0,Null,Net),If(length(Datemnf) < 3,' ',Datemnf),Concat(Exvessel,'-',ExVoy),
    UPPER(interchange.Consignee),Date(EIRIN.DateIn),
    EIRIN.PrincipleCode,EIRIN.Remark,EIRIN.Grade,EIRIN.IntNo 
    From EIRIN 
    left Join Interchange On Interchange.Nomor = EIRIN.IntNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo and 
    InterchangeContainer.ContNo = EIRIn.ContNo 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIn.ContNo 
    left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.contno = eirin.contno and InterchangeDocpaycontainer.intno=EIRIN.intNo 
    left Join interchangeDocpaydetails on InterchangeDocpaydetails.nomor = InterchangeDocpaycontainer.nomor 
    and InterchangeDocpaydetails.Size = ContainerDetails.Size 
    where EIRIn.PrincipleCode = '"""+ client +"""' and EIRIN.DateIn >= '"""+ start_date +""" 00:00:01' and EIRIN.DateIn <= '"""+ end_date +""" 23:59:59' 
    GROUP BY EIRIN.NOMOR order by EIRIN.DateIn"""

    sql2 = """Select EIRIN.ContNo,ContainerDetails.size,ContainerDetails.type,contCondition,
    Payload,If(Net=0,Null,Net),If(length(Datemnf) < 3,' ',Datemnf),Concat(Exvessel,'-',ExVoy),
    interchange.Consignee,Date(EIRIN.DateIn),Time(EIRIN.DateIn),Date(EIRIN.DateIn),Date(EIRIN.DateIn),
    EIRIN.PrincipleCode,VN,VN,VN,EIRIN.Remark,EIRIN.Grade,EIRIN.IntNo 
    From EIRIN 
    left Join Interchange On Interchange.Nomor = EIRIN.IntNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo and 
    InterchangeContainer.ContNo = EIRIn.ContNo 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIn.ContNo 
    left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.contno = eirin.contno and InterchangeDocpaycontainer.intno=EIRIN.intNo 
    left Join interchangeDocpaydetails on InterchangeDocpaydetails.nomor = InterchangeDocpaycontainer.nomor 
    and InterchangeDocpaydetails.Size = ContainerDetails.Size 
    where EIRIn.PrincipleCode = '"""+ client +"""' and EIRIN.DateIn >= '"""+ start_date +""" 00:00:01' and EIRIN.DateIn <= '"""+ end_date +""" 23:59:59' 
    GROUP BY EIRIN.NOMOR order by EIRIN.DateIn"""

    '''
    sql3 = """Select EIRIN.ContNo,ContainerDetails.size,ContainerDetails.type,Interchange.Consignee,
    EIRIN.IntNo,Exvessel,ExVoy,Date(EIRIN.DateIn),EIRIN.Remark,VN,contCondition,
    InterchangeContainer.CleaningType,EIRIN.Nomor,If(length(Datemnf) < 3,' ',Datemnf),
    If(MGW=0,Null,MGW),Payload,If(Net=0,Null,Net),EIRIN.DateOutPort,
    If(Exrepo = 0,'Ex. Import','Ex. Repo'),EIRIN.PrincipleCode 
    From EIRIN 
    left Join Interchange On Interchange.Nomor = EIRIN.IntNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo and 
    InterchangeContainer.ContNo = EIRIn.ContNo 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIn.ContNo 
    left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.contno = eirin.contno and InterchangeDocpaycontainer.intno=EIRIN.intNo 
    left Join interchangeDocpaydetails on InterchangeDocpaydetails.nomor = InterchangeDocpaycontainer.nomor 
    and InterchangeDocpaydetails.Size = ContainerDetails.Size 
    where EIRIn.PrincipleCode = '"""+ client +"""' and EIRIN.DateIn >= '2021-07-12 00:00:01' and EIRIN.DateIn <= '2021-07-18 23:59:59' 
    GROUP BY EIRIN.NOMOR order by EIRIN.DateIn"""

    sql4 = """Select EIRIN.ContNo,ContainerDetails.size,ContainerDetails.type,Interchange.Consignee,
    EIRIN.IntNo,Exvessel,ExVoy,Date(EIRIN.DateIn),EIRIN.Remark,VN,contCondition,
    InterchangeContainer.CleaningType,EIRIN.Nomor,If(length(Datemnf) < 3,' ',Datemnf),
    If(MGW=0,Null,MGW),Payload,If(Net=0,Null,Net),EIRIN.DateOutPort,
    If(Exrepo = 0,'Ex. Import','Ex. Repo'),EIRIN.PrincipleCode 
    From EIRIN 
    left Join Interchange On Interchange.Nomor = EIRIN.IntNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo and 
    InterchangeContainer.ContNo = EIRIn.ContNo 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIn.ContNo 
    left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.contno = eirin.contno and InterchangeDocpaycontainer.intno=EIRIN.intNo 
    left Join interchangeDocpaydetails on InterchangeDocpaydetails.nomor = InterchangeDocpaycontainer.nomor 
    and InterchangeDocpaydetails.Size = ContainerDetails.Size 
    where EIRIn.PrincipleCode = '"""+ client +"""' and EIRIN.DateIn >= '2021-07-12 00:00:01' and EIRIN.DateIn <= '2021-07-18 23:59:59' 
    GROUP BY EIRIN.NOMOR order by EIRIN.DateIn"""

    sql5 = """Select EIRIN.ContNo,ContainerDetails.size,ContainerDetails.type,Interchange.Consignee,
    EIRIN.IntNo,Exvessel,ExVoy,Date(EIRIN.DateIn),EIRIN.Remark,VN,contCondition,
    InterchangeContainer.CleaningType,EIRIN.Nomor,If(length(Datemnf) < 3,' ',Datemnf),
    If(MGW=0,Null,MGW),Payload,If(Net=0,Null,Net),EIRIN.DateOutPort,
    If(Exrepo = 0,'Ex. Import','Ex. Repo'),EIRIN.PrincipleCode 
    From EIRIN 
    left Join Interchange On Interchange.Nomor = EIRIN.IntNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo and 
    InterchangeContainer.ContNo = EIRIn.ContNo 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIn.ContNo 
    left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.contno = eirin.contno and InterchangeDocpaycontainer.intno=EIRIN.intNo 
    left Join interchangeDocpaydetails on InterchangeDocpaydetails.nomor = InterchangeDocpaycontainer.nomor 
    and InterchangeDocpaydetails.Size = ContainerDetails.Size 
    where EIRIn.PrincipleCode = 'CMA' and EIRIN.DateIn >= '2021-07-12 00:00:01' and EIRIN.DateIn <= '2021-07-18 23:59:59' 
    GROUP BY EIRIN.NOMOR order by EIRIN.DateIn"""

    sql6 = """Select EIRIN.ContNo,ContainerDetails.size,ContainerDetails.type,Interchange.Consignee,
    EIRIN.IntNo,Exvessel,ExVoy,Date(EIRIN.DateIn),EIRIN.Remark,VN,contCondition,
    InterchangeContainer.CleaningType,EIRIN.Nomor,If(length(Datemnf) < 3,' ',Datemnf),
    If(MGW=0,Null,MGW),Payload,If(Net=0,Null,Net),EIRIN.DateOutPort,
    If(Exrepo = 0,'Ex. Import','Ex. Repo'),EIRIN.PrincipleCode 
    From EIRIN 
    left Join Interchange On Interchange.Nomor = EIRIN.IntNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo and 
    InterchangeContainer.ContNo = EIRIn.ContNo 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIn.ContNo 
    left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.contno = eirin.contno and InterchangeDocpaycontainer.intno=EIRIN.intNo 
    left Join interchangeDocpaydetails on InterchangeDocpaydetails.nomor = InterchangeDocpaycontainer.nomor 
    and InterchangeDocpaydetails.Size = ContainerDetails.Size 
    where EIRIn.PrincipleCode = 'CMA' and EIRIN.DateIn >= '2021-07-12 00:00:01' and EIRIN.DateIn <= '2021-07-18 23:59:59' 
    GROUP BY EIRIN.NOMOR order by EIRIN.DateIn"""
    '''
    excel_filename = ''+ client +' - '+ start_date +' '+ end_date +' - REPORT.xlsx'
    # Writing Database into Excel
    writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
    # Return sql query result
    df1 = pd.read_sql_query(sql1, db)
    df2 = pd.read_sql_query(sql2, db)
    '''
    df3 = pd.read_sql_query(sql3, db)
    df4 = pd.read_sql_query(sql4, db)
    df5 = pd.read_sql_query(sql5, db)
    df6 = pd.read_sql_query(sql6, db)

    '''
    df1.index += 1
    df2.index += 1
    header = [
    "Container No","Size","Type","Condition","Payload","Tare","Date Mnf","Ex Vessel Voy",
    "Customer","Date In","Principal",
    "Remarks Restriction","Grade","B/L NO"
    ]

    df1.to_excel(writer, sheet_name='DAILY STOCK POSITION', startrow=5, header=header)
    df2.to_excel(writer, sheet_name='STOCK LIST', startrow=5)
    '''
    df3.to_excel(writer, sheet_name='MOV IN', startrow=5)
    df4.to_excel(writer, sheet_name='MOV OUT',startrow=5)
    df5.to_excel(writer, sheet_name='DAMAGE', startrow=5)
    df6.to_excel(writer, sheet_name='FINISHED REPAIR',startrow=5)
    
    worksheet2 = writer.sheets['STOCK LIST']
    '''
    #auto_adjust_xlsx_column_width(df2, writer, sheet_name='STOCK LIST', margin=0)
    # Formatting Excel File
    workbook = writer.book
    worksheet = writer.sheets['DAILY STOCK POSITION']
    worksheet2 = writer.sheets['STOCK LIST']
    header_format = workbook.add_format({
        'bg_color': '#E7E6E6',
        'bold': False,
        'align': 'center',
    })
    column_number_format = workbook.add_format({
        'bold': False,
        'align': 'center',
    })

    merge_info = workbook.add_format({
        'bold': False
    })
    merge_info1 = workbook.add_format({
        'bold': True
    })
    merge_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'bold': True
    })
    merge_format_number = workbook.add_format({
        'align': 'center',
        'bold': False
    })


    align_center = workbook.add_format({
        'align': 'center'
    })

    #worksheet.set_column("E:E", 16, cell_format=yellow)
    #for idx in enumerate(df1.columns):
    #    worksheet.write(6, idx, None, header_format)
    # WORKSHEET 1 FORMAT
    # Merge necessary column
    worksheet.merge_range('A1:E1', 'Customer : PT. CONTAINER MARITIME ACTIVITES', merge_info)
    worksheet.merge_range('A2:E2', 'From date : '+ start_date +' 00:00:01', merge_info)
    worksheet.merge_range('A3:E3', 'To date : '+ end_date +' 23:59:59', merge_info)
    #worksheet.merge_range('G6:H6', 'Ex - Vessel', merge_format)
    #worksheet.merge_range('O6:R6', 'Cont Detail', merge_format)
    
    # Add Number Iteration
    worksheet.write('A6:A6', 'No.', merge_format_number)

    # Format column Text Align
    worksheet.set_column('B:I', None, align_center) 
    worksheet.set_column('K:L', None, align_center)
    #worksheet2.set_column('I:I', None, align_center)  
    worksheet.set_column('N:O', None, align_center)
    worksheet.set_column(0, 0, 5)

    '''
    for column in df1:
        column_length = max(df1[column].astype(str).map(len).max(), len(column))
        col_idx = df1.columns.get_loc(column)
        worksheet2.set_column(col_idx+5, col_idx+5, column_length)
    '''
    # Auto-adjust column width
    for idx, col in enumerate(df1.columns):  # loop through all columns
        series = df1[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 2  # adding a little extra space
        worksheet.set_column(idx+1, idx+1, max_len)
    #auto_adjust_xlsx_column_width(df1, writer, sheet_name="DAILY STOCK POSITION", margin=0)

    # Un-bold the number iteration
    for row_num, value in enumerate(df1.index.get_level_values(level=0)):
        worksheet.write(row_num+6, 0, value, column_number_format)
    

    # Adding border to data    
    #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
    #worksheet.conditional_format(xlsxwriter.utility.xl_range(6, 0, len(df1.index), len(df1.columns)), {'type': 'no_errors', 'format': border_fmt})
    
    # WORKSHEET 2 FORMAT
    # Merge necessary column
    worksheet2.merge_range('A1:E1', 'Customer : PT. CONTAINER MARITIME ACTIVITES', merge_info1)
    worksheet2.merge_range('A2:E2', 'From date :  From date : 2021-07-12 00:00:01', merge_info)
    worksheet2.merge_range('A3:E3', 'To date : 2021-07-18 23:59:59', merge_info)
    
    
    
    # Add Number Iteration
    worksheet2.write('A6', 'No.', merge_format)
    # Auto-adjust column width

    

    # Un-bold the number iteration
    for row_num, value in enumerate(df2.index.get_level_values(level=0)):
        worksheet2.write(row_num+6, 0, value, header_format)
    # Format column Text Align
    worksheet2.set_column('B:I', None, align_center) 
    worksheet2.set_column('K:L', None, align_center)
    worksheet2.set_column('I:I', None, align_center)  
    worksheet2.set_column('N:O', None, align_center)
    worksheet2.set_column(0, 0, 5)

    
    

    # Adding border to data    
    #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
    #worksheet2.conditional_format(xlsxwriter.utility.xl_range(6, 0, len(df2.index), len(df2.columns)), {'type': 'no_errors', 'format': border_fmt})

    writer.save()

    fromaddr = "linibelajar@gmail.com"
    #toaddr = "oktoberlin@gmail.com"
    recipients = email
    msg = MIMEMultipart()

    msg['From'] = fromaddr
    msg['To'] = ", ".join(recipients)
    msg['Subject'] = "Hasil Generate Excel"

    body = "Berikut adalah hasil generate excel"

    msg.attach(MIMEText(body, 'plain'))

    filename = excel_filename
    attachment = open(filename, "rb")

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(fromaddr, "hyrxthrbsiicfljk")
    text = msg.as_string()
    server.sendmail(fromaddr, recipients, text)
    server.quit()

@shared_task
def send_mail_task_daily():
    tasks = generate_excel.objects.all()

    for task in tasks:
        client = task.client
        email=task.email
        start_date = task.start_date
        end_date = task.end_date
        # connecting server database to python
        db = pymysql.connect(user="admin21",password="adminsby21",host="149.129.243.10",port=3307,database="mydepo")
        sql1 = """Select EIRIN.ContNo,ContainerDetails.size,ContainerDetails.type,contCondition,
        Payload,If(Net=0,Null,Net),If(length(Datemnf) < 3,' ',Datemnf),Concat(Exvessel,'-',ExVoy),
        UPPER(interchange.Consignee),Date(EIRIN.DateIn),
        EIRIN.PrincipleCode,EIRIN.Remark,EIRIN.Grade,EIRIN.IntNo 
        From EIRIN 
        left Join Interchange On Interchange.Nomor = EIRIN.IntNo 
        left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo and 
        InterchangeContainer.ContNo = EIRIn.ContNo 
        Left join ContainerDetails On ContainerDetails.ContNo = EIRIn.ContNo 
        left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.contno = eirin.contno and InterchangeDocpaycontainer.intno=EIRIN.intNo 
        left Join interchangeDocpaydetails on InterchangeDocpaydetails.nomor = InterchangeDocpaycontainer.nomor 
        and InterchangeDocpaydetails.Size = ContainerDetails.Size 
        where EIRIn.PrincipleCode = '"""+ client +"""' and EIRIN.DateIn >= '"""+ start_date +""" 00:00:01' and EIRIN.DateIn <= '"""+ end_date +""" 23:59:59' 
        GROUP BY EIRIN.NOMOR order by EIRIN.DateIn"""

        sql2 = """Select EIRIN.ContNo,ContainerDetails.size,ContainerDetails.type,contCondition,
        Payload,If(Net=0,Null,Net),If(length(Datemnf) < 3,' ',Datemnf),Concat(Exvessel,'-',ExVoy),
        interchange.Consignee,Date(EIRIN.DateIn),Time(EIRIN.DateIn),Date(EIRIN.DateIn),Date(EIRIN.DateIn),
        EIRIN.PrincipleCode,VN,VN,VN,EIRIN.Remark,EIRIN.Grade,EIRIN.IntNo 
        From EIRIN 
        left Join Interchange On Interchange.Nomor = EIRIN.IntNo 
        left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo and 
        InterchangeContainer.ContNo = EIRIn.ContNo 
        Left join ContainerDetails On ContainerDetails.ContNo = EIRIn.ContNo 
        left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.contno = eirin.contno and InterchangeDocpaycontainer.intno=EIRIN.intNo 
        left Join interchangeDocpaydetails on InterchangeDocpaydetails.nomor = InterchangeDocpaycontainer.nomor 
        and InterchangeDocpaydetails.Size = ContainerDetails.Size 
        where EIRIn.PrincipleCode = '"""+ client +"""' and EIRIN.DateIn >= '"""+ start_date +""" 00:00:01' and EIRIN.DateIn <= '"""+ end_date +""" 23:59:59' 
        GROUP BY EIRIN.NOMOR order by EIRIN.DateIn"""

        '''
        sql3 = """Select EIRIN.ContNo,ContainerDetails.size,ContainerDetails.type,Interchange.Consignee,
        EIRIN.IntNo,Exvessel,ExVoy,Date(EIRIN.DateIn),EIRIN.Remark,VN,contCondition,
        InterchangeContainer.CleaningType,EIRIN.Nomor,If(length(Datemnf) < 3,' ',Datemnf),
        If(MGW=0,Null,MGW),Payload,If(Net=0,Null,Net),EIRIN.DateOutPort,
        If(Exrepo = 0,'Ex. Import','Ex. Repo'),EIRIN.PrincipleCode 
        From EIRIN 
        left Join Interchange On Interchange.Nomor = EIRIN.IntNo 
        left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo and 
        InterchangeContainer.ContNo = EIRIn.ContNo 
        Left join ContainerDetails On ContainerDetails.ContNo = EIRIn.ContNo 
        left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.contno = eirin.contno and InterchangeDocpaycontainer.intno=EIRIN.intNo 
        left Join interchangeDocpaydetails on InterchangeDocpaydetails.nomor = InterchangeDocpaycontainer.nomor 
        and InterchangeDocpaydetails.Size = ContainerDetails.Size 
        where EIRIn.PrincipleCode = '"""+ client +"""' and EIRIN.DateIn >= '2021-07-12 00:00:01' and EIRIN.DateIn <= '2021-07-18 23:59:59' 
        GROUP BY EIRIN.NOMOR order by EIRIN.DateIn"""

        sql4 = """Select EIRIN.ContNo,ContainerDetails.size,ContainerDetails.type,Interchange.Consignee,
        EIRIN.IntNo,Exvessel,ExVoy,Date(EIRIN.DateIn),EIRIN.Remark,VN,contCondition,
        InterchangeContainer.CleaningType,EIRIN.Nomor,If(length(Datemnf) < 3,' ',Datemnf),
        If(MGW=0,Null,MGW),Payload,If(Net=0,Null,Net),EIRIN.DateOutPort,
        If(Exrepo = 0,'Ex. Import','Ex. Repo'),EIRIN.PrincipleCode 
        From EIRIN 
        left Join Interchange On Interchange.Nomor = EIRIN.IntNo 
        left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo and 
        InterchangeContainer.ContNo = EIRIn.ContNo 
        Left join ContainerDetails On ContainerDetails.ContNo = EIRIn.ContNo 
        left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.contno = eirin.contno and InterchangeDocpaycontainer.intno=EIRIN.intNo 
        left Join interchangeDocpaydetails on InterchangeDocpaydetails.nomor = InterchangeDocpaycontainer.nomor 
        and InterchangeDocpaydetails.Size = ContainerDetails.Size 
        where EIRIn.PrincipleCode = '"""+ client +"""' and EIRIN.DateIn >= '2021-07-12 00:00:01' and EIRIN.DateIn <= '2021-07-18 23:59:59' 
        GROUP BY EIRIN.NOMOR order by EIRIN.DateIn"""

        sql5 = """Select EIRIN.ContNo,ContainerDetails.size,ContainerDetails.type,Interchange.Consignee,
        EIRIN.IntNo,Exvessel,ExVoy,Date(EIRIN.DateIn),EIRIN.Remark,VN,contCondition,
        InterchangeContainer.CleaningType,EIRIN.Nomor,If(length(Datemnf) < 3,' ',Datemnf),
        If(MGW=0,Null,MGW),Payload,If(Net=0,Null,Net),EIRIN.DateOutPort,
        If(Exrepo = 0,'Ex. Import','Ex. Repo'),EIRIN.PrincipleCode 
        From EIRIN 
        left Join Interchange On Interchange.Nomor = EIRIN.IntNo 
        left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo and 
        InterchangeContainer.ContNo = EIRIn.ContNo 
        Left join ContainerDetails On ContainerDetails.ContNo = EIRIn.ContNo 
        left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.contno = eirin.contno and InterchangeDocpaycontainer.intno=EIRIN.intNo 
        left Join interchangeDocpaydetails on InterchangeDocpaydetails.nomor = InterchangeDocpaycontainer.nomor 
        and InterchangeDocpaydetails.Size = ContainerDetails.Size 
        where EIRIn.PrincipleCode = 'CMA' and EIRIN.DateIn >= '2021-07-12 00:00:01' and EIRIN.DateIn <= '2021-07-18 23:59:59' 
        GROUP BY EIRIN.NOMOR order by EIRIN.DateIn"""

        sql6 = """Select EIRIN.ContNo,ContainerDetails.size,ContainerDetails.type,Interchange.Consignee,
        EIRIN.IntNo,Exvessel,ExVoy,Date(EIRIN.DateIn),EIRIN.Remark,VN,contCondition,
        InterchangeContainer.CleaningType,EIRIN.Nomor,If(length(Datemnf) < 3,' ',Datemnf),
        If(MGW=0,Null,MGW),Payload,If(Net=0,Null,Net),EIRIN.DateOutPort,
        If(Exrepo = 0,'Ex. Import','Ex. Repo'),EIRIN.PrincipleCode 
        From EIRIN 
        left Join Interchange On Interchange.Nomor = EIRIN.IntNo 
        left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo and 
        InterchangeContainer.ContNo = EIRIn.ContNo 
        Left join ContainerDetails On ContainerDetails.ContNo = EIRIn.ContNo 
        left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.contno = eirin.contno and InterchangeDocpaycontainer.intno=EIRIN.intNo 
        left Join interchangeDocpaydetails on InterchangeDocpaydetails.nomor = InterchangeDocpaycontainer.nomor 
        and InterchangeDocpaydetails.Size = ContainerDetails.Size 
        where EIRIn.PrincipleCode = 'CMA' and EIRIN.DateIn >= '2021-07-12 00:00:01' and EIRIN.DateIn <= '2021-07-18 23:59:59' 
        GROUP BY EIRIN.NOMOR order by EIRIN.DateIn"""
        '''
        excel_filename = ''+ client +' - '+ start_date +' '+ end_date +' - REPORT.xlsx'
        # Writing Database into Excel
        writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
        # Return sql query result
        df1 = pd.read_sql_query(sql1, db)
        df2 = pd.read_sql_query(sql2, db)
        '''
        df3 = pd.read_sql_query(sql3, db)
        df4 = pd.read_sql_query(sql4, db)
        df5 = pd.read_sql_query(sql5, db)
        df6 = pd.read_sql_query(sql6, db)

        '''
        df1.index += 1
        df2.index += 1
        header = [
        "Container No","Size","Type","Condition","Payload","Tare","Date Mnf","Ex Vessel Voy",
        "Customer","Date In","Principal",
        "Remarks Restriction","Grade","B/L NO"
        ]

        df1.to_excel(writer, sheet_name='DAILY STOCK POSITION', startrow=5, header=header)
        df2.to_excel(writer, sheet_name='STOCK LIST', startrow=5)
        '''
        df3.to_excel(writer, sheet_name='MOV IN', startrow=5)
        df4.to_excel(writer, sheet_name='MOV OUT',startrow=5)
        df5.to_excel(writer, sheet_name='DAMAGE', startrow=5)
        df6.to_excel(writer, sheet_name='FINISHED REPAIR',startrow=5)
        
        worksheet2 = writer.sheets['STOCK LIST']
        '''
        #auto_adjust_xlsx_column_width(df2, writer, sheet_name='STOCK LIST', margin=0)
        # Formatting Excel File
        workbook = writer.book
        worksheet = writer.sheets['DAILY STOCK POSITION']
        worksheet2 = writer.sheets['STOCK LIST']
        header_format = workbook.add_format({
            'bg_color': '#E7E6E6',
            'bold': False,
            'align': 'center',
        })
        column_number_format = workbook.add_format({
            'bold': False,
            'align': 'center',
        })

        merge_info = workbook.add_format({
            'bold': False
        })
        merge_info1 = workbook.add_format({
            'bold': True
        })
        merge_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'bold': True
        })
        merge_format_number = workbook.add_format({
            'align': 'center',
            'bold': False
        })


        align_center = workbook.add_format({
            'align': 'center'
        })

        #worksheet.set_column("E:E", 16, cell_format=yellow)
        #for idx in enumerate(df1.columns):
        #    worksheet.write(6, idx, None, header_format)
        # WORKSHEET 1 FORMAT
        # Merge necessary column
        worksheet.merge_range('A1:E1', 'Customer : PT. CONTAINER MARITIME ACTIVITES', merge_info)
        worksheet.merge_range('A2:E2', 'From date : '+ start_date +' 00:00:01', merge_info)
        worksheet.merge_range('A3:E3', 'To date : '+ end_date +' 23:59:59', merge_info)
        #worksheet.merge_range('G6:H6', 'Ex - Vessel', merge_format)
        #worksheet.merge_range('O6:R6', 'Cont Detail', merge_format)
        
        # Add Number Iteration
        worksheet.write('A6:A6', 'No.', merge_format_number)

        # Format column Text Align
        worksheet.set_column('B:I', None, align_center) 
        worksheet.set_column('K:L', None, align_center)
        #worksheet2.set_column('I:I', None, align_center)  
        worksheet.set_column('N:O', None, align_center)
        worksheet.set_column(0, 0, 5)

        '''
        for column in df1:
            column_length = max(df1[column].astype(str).map(len).max(), len(column))
            col_idx = df1.columns.get_loc(column)
            worksheet2.set_column(col_idx+5, col_idx+5, column_length)
        '''
        # Auto-adjust column width
        for idx, col in enumerate(df1.columns):  # loop through all columns
            series = df1[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
                )) + 2  # adding a little extra space
            worksheet.set_column(idx+1, idx+1, max_len)
        #auto_adjust_xlsx_column_width(df1, writer, sheet_name="DAILY STOCK POSITION", margin=0)

        # Un-bold the number iteration
        for row_num, value in enumerate(df1.index.get_level_values(level=0)):
            worksheet.write(row_num+6, 0, value, column_number_format)
        

        # Adding border to data    
        #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        #worksheet.conditional_format(xlsxwriter.utility.xl_range(6, 0, len(df1.index), len(df1.columns)), {'type': 'no_errors', 'format': border_fmt})
        
        # WORKSHEET 2 FORMAT
        # Merge necessary column
        worksheet2.merge_range('A1:E1', 'Customer : PT. CONTAINER MARITIME ACTIVITES', merge_info1)
        worksheet2.merge_range('A2:E2', 'From date :  From date : 2021-07-12 00:00:01', merge_info)
        worksheet2.merge_range('A3:E3', 'To date : 2021-07-18 23:59:59', merge_info)
        
        
        
        # Add Number Iteration
        worksheet2.write('A6', 'No.', merge_format)
        # Auto-adjust column width

        

        # Un-bold the number iteration
        for row_num, value in enumerate(df2.index.get_level_values(level=0)):
            worksheet2.write(row_num+6, 0, value, header_format)
        # Format column Text Align
        worksheet2.set_column('B:I', None, align_center) 
        worksheet2.set_column('K:L', None, align_center)
        worksheet2.set_column('I:I', None, align_center)  
        worksheet2.set_column('N:O', None, align_center)
        worksheet2.set_column(0, 0, 5)

        
        

        # Adding border to data    
        #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        #worksheet2.conditional_format(xlsxwriter.utility.xl_range(6, 0, len(df2.index), len(df2.columns)), {'type': 'no_errors', 'format': border_fmt})

        writer.save()

        fromaddr = "linibelajar@gmail.com"
        #toaddr = "oktoberlin@gmail.com"
        recipients = email
        msg = MIMEMultipart()

        msg['From'] = fromaddr
        msg['To'] = ", ".join(recipients)
        msg['Subject'] = "Hasil Generate Excel"

        body = "Berikut adalah hasil generate excel"

        msg.attach(MIMEText(body, 'plain'))

        filename = excel_filename
        attachment = open(filename, "rb")

        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(fromaddr, "hyrxthrbsiicfljk")
        text = msg.as_string()
        server.sendmail(fromaddr, recipients, text)
        server.quit()