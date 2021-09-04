# task.py, this file can be present in any of the apps.
from io import StringIO
from celery.schedules import crontab

from celery import shared_task
import pandas as pd
import xlsxwriter
import pymysql
import smtplib
from smtplib import *
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from .models import generate_excel
from datetime import datetime, time, timedelta
import pytz

@shared_task
def mysql_to_excel(time_now_filename,Client,time_now_email_subject,time_from):
    
    '''
    if f'{time_now_email_subject} 16:00:00' <= time_now <= f'{time_now_email_subject} 16:01:00' and sent==False:
        time_from= f'{time_from} 09:00:01'
        sent=True
        print('good')

    if '2021-09-04 00:00:00' <= time_now <=  '2021-09-04 00:01:00' and sent==False:
        time_from = time_yesterday
        sent=True
        print('really?')
    '''
    # File name

    #time_from = '2021-09-03 00:00:01'
    #time_now = datetime.now(pytz.timezone('Asia/Jakarta')).strftime("%Y-%m-%d %H:%M:%S")
    print(time_now)
    print(time_from)
    '''
    time_yesterday = (datetime.now(pytz.timezone('Asia/Jakarta'))-timedelta(1)).strftime("%Y-%m-%d %H:%M:%S")
    time_now_email_subject = datetime.now(pytz.timezone('Asia/Jakarta')).strftime("%Y-%m-%d")
    time_now = '2021-09-03 16:00:10'
    #time_from = f'{time_now_email_subject} 00:00:01'                                   

    if f'{time_now_email_subject} 09:00:00' <= time_now <= f'{time_now_email_subject} 15:59:59':
      time_from = f'{time_now_email_subject} 00:00:01'
    if f'{time_now_email_subject} 16:00:00' <= time_now <= f'{time_now_email_subject} 23:59:59':
      time_from = f'{time_now_email_subject} 09:00:01'
      print(time_from)
    if f'{time_now_email_subject} 00:00:00' <= time_now <= f'{time_now_email_subject} 08:59:59':
      time_from = time_yesterday

    '''
    #if f'{time_now_email_subject} 16:00:01' <= time_now <= f'{time_now_email_subject} 00:00:01':
    #  time_from = time_yesterday
    print("Please don't close this task scheduler running")
    #print(time_from)
    
    Principle_Code = 'PNP'
    # connecting server database Jakarta to pythons
    db = pymysql.connect(user="adminsby",password="adminsby21",host="103.105.68.214",port=3333,database="mydepo")
    sql1 = """
    SELECT * FROM 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'FR' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'FR'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'FR' AND 
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""')
    ) AS A11 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'FR' AND  ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'FR'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'FR' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A12 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'FR' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'FR' 
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'FR' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A13 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'FR' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS B1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'FR' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS C1
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'FR' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS D1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND
    InterchangeContainer.CleaningType IN ('%DW%','%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND
    InterchangeContainer.CleaningType IN ('%DW%','%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND
    InterchangeContainer.CleaningType IN ('%DW%','%CW%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'FR' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS E1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'FR' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS F1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'FR' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS G1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' THEN 1 END)
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    WHERE RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'FR' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS H1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' THEN 1 END)
    FROM EIROUT 
    Left Join EIRIN on EIRIN.Nomor = EIROUT.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE ContainerDetails.type = 'FR' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""'
    AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""'
    ) AS I1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'FR' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS J1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END)  
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'FR' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS K1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'FR' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS L1 
    UNION ALL 
    SELECT * FROM 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'GP' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'GP'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'GP' AND 
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""')
    ) AS A11 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'GP' AND  ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'GP'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'GP' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A12 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'GP' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'GP' 
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'GP' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A13 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'GP' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS B1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'GP' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS C1
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'GP' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS D1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'GP' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS E1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'GP' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS F1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'GP' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS G1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END)
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    WHERE RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'GP' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS H1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END)
    FROM EIROUT 
    Left Join EIRIN on EIRIN.Nomor = EIROUT.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE ContainerDetails.type = 'GP' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""'
    AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""'
    ) AS I1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'GP' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS J1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END)  
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'GP' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS K1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'GP' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS L1 
    UNION ALL 
    SELECT * FROM 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'GOH' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'GOH'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'GOH' AND 
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""')
    ) AS A11 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'GOH' AND  ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'GOH'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'GOH' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A12 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'GOH' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'GOH' 
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'GOH' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A13 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'GOH' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS B1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'GOH' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS C1
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'GOH' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS D1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'GOH' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS E1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'GOH' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS F1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'GOH' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS G1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END)
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    WHERE RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'GOH' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS H1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END)
    FROM EIROUT 
    Left Join EIRIN on EIRIN.Nomor = EIROUT.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE ContainerDetails.type = 'GOH' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""'
    AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""'
    ) AS I1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'GOH' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS J1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END)  
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'GOH' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS K1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'GOH' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS L1 
    UNION ALL 
    SELECT * FROM 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'HC' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'HC'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'HC' AND 
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""')
    ) AS A11 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'HC' AND  ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'HC'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'HC' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A12 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'HC' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'HC' 
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'HC' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A13 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'HC' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS B1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'HC' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS C1
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'HC' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS D1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'HC' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS E1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'HC' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS F1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'HC' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS G1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END)
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    WHERE RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'HC' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS H1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END)
    FROM EIROUT 
    Left Join EIRIN on EIRIN.Nomor = EIROUT.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE ContainerDetails.type = 'HC' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""'
    AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""'
    ) AS I1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'HC' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS J1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END)  
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'HC' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS K1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'HC' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS L1
    UNION ALL 
    SELECT * FROM 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'OT' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'OT'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'OT' AND 
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""')
    ) AS A11 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'OT' AND  ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'OT'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'OT' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A12 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'OT' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'OT' 
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'OT' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A13 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'OT' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS B1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'OT' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS C1
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'OT' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS D1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'OT' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS E1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'OT' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS F1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'OT' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS G1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END)
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    WHERE RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'OT' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS H1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END)
    FROM EIROUT 
    Left Join EIRIN on EIRIN.Nomor = EIROUT.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE ContainerDetails.type = 'OT' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""'
    AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""'
    ) AS I1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'OT' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS J1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END)  
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'OT' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS K1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'OT' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS L1 
    UNION ALL 
    SELECT * FROM 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'RF' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'RF'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'RF' AND 
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""')
    ) AS A11 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'RF' AND  ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'RF'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'RF' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A12 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'RF' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'RF' 
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'RF' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A13 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'RF' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS B1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'RF' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS C1
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'RF' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS D1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'RF' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS E1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'RF' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS F1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'RF' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS G1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END)
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    WHERE RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'RF' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS H1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END)
    FROM EIROUT 
    Left Join EIRIN on EIRIN.Nomor = EIROUT.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE ContainerDetails.type = 'RF' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""'
    AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""'
    ) AS I1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'RF' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS J1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END)  
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'RF' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS K1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'RF' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS L1 
    UNION ALL 
    SELECT * FROM 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'RH' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'RH'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND ContainerDetails.type = 'RH' AND 
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""')
    ) AS A11 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'RH' AND  ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'RH'
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND ContainerDetails.type = 'RH' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A12 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'RH' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'RH' 
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND ContainerDetails.type = 'RH' AND
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A13 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'RH' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS B1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'RH' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS C1
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'RH' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS D1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'RH' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS E1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'RH' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS F1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE ContainerDetails.type = 'RH' AND EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS G1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END)
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    WHERE RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'RH' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS H1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (EIROUT.contCondition LIKE '%AV%' OR EIROUT.contCondition LIKE '%DMG%') THEN 1 END)
    FROM EIROUT 
    Left Join EIRIN on EIRIN.Nomor = EIROUT.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE ContainerDetails.type = 'RH' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""'
    AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""'
    ) AS I1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%') THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'RH' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS J1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END)  
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'RH' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS K1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerDetails.type = 'RH' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS L1 
    UNION ALL 
    SELECT * FROM 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '20' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' 
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '20' AND 
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""')
    ) AS A11 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '40' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' 
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '40' AND 
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A12 
    join 
    (SELECT 
    (SELECT COUNT(*) FROM ContainerStock 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    where  EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""' AND ContainerDetails.size = '45' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""') +
    (SELECT COUNT(*) FROM EIROUT 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' 
    AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""')-
    (SELECT COUNT(*) FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""' AND ContainerDetails.size = '45' AND 
    EIRIN.DateIn >= '"""+time_from+"""' 
    AND EIRIN.DateIn <= '"""+time_now+"""') 
    ) AS A13 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS B1
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%AV%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS C1
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND EIRIN.contCondition LIKE '%DMG%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS D1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND 
    (InterchangeContainer.CleaningType LIKE '%DW%' OR InterchangeContainer.CleaningType LIKE '%CW%') THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS E1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%WW%' THEN 1 END)
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS F1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND InterchangeContainer.CleaningType LIKE '%SW%' THEN 1 END) 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    WHERE EIRIN.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS G1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' THEN 1 END)
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    WHERE RepairContainer.Repaired = 'Yes' AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS H1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' THEN 1 END)
    FROM EIROUT 
    Left Join EIRIN on EIRIN.Nomor = EIROUT.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    WHERE EIROUT.PrincipleCode = '"""+Principle_Code+"""'
    AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""'
    ) AS I1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' THEN 1 END)
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS J1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%AV%' THEN 1 END)  
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS K1 
    JOIN 
    (SELECT 
    COUNT(CASE WHEN ContainerDetails.size = '20' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '40' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END),
    COUNT(CASE WHEN ContainerDetails.size = '45' AND ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    WHERE ContainerStock.PrincipleCode = '"""+Principle_Code+"""'
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'
    ) AS L1
    """

    sql2 = """Select 
    ContainerStock.ContNo as 'Container No',ContainerDetails.Size as 'Size',ContainerDetails.Type as 'Type',
    ContainerStock.ContCondition as 'Condition',ContainerDetails.Payload as 'Payload',
    If(ContainerDetails.Net=0,Null,Net) as 'Tare',If(length(ContainerDetails.Datemnf) < 3,' ',
    ContainerDetails.Datemnf) as 'Date Mnf',Concat(Interchange.Exvessel,'-',Interchange.ExVoy) as 'Ex Vessel Voy',
    UPPER(Interchange.Consignee) as 'Customer',EIRIN.DateIn as 'Date IN',ContainerStock.PrincipleCode as Principle,
    BlockContainer.Remark as 'Remarks IN',EIRIN.Grade as Grade,EIRIN.IntNo as 'B/L NO' 
    From ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join BlockContainer On BlockContainer.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left Join Interchange On Interchange.Nomor = ContainerStock.IntNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = ContainerStock.IntNo AND 
    InterchangeContainer.ContNo = ContainerStock.ContNo 
    left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.ContNo = ContainerStock.ContNo 
    AND InterchangeDocpaycontainer.IntNo = ContainerStock.IntNo 
    left Join interchangeDocpaydetails on InterchangeDocpaydetails.Nomor = InterchangeDocpaycontainer.Nomor 
    AND InterchangeDocpaydetails.Size = ContainerDetails.Size 
    where ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    GROUP BY ContainerStock.ContNo order by EIRIN.DateIn"""

    sql2_summary = """
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' 
    OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'GP' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' 
    OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'HC' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' 
    OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'OT' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' 
    OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'FR' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' 
    OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'RF' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' 
    OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'TK' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' 
    OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'GP' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' 
    OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'HC' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' 
    OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'OT' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'   
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' 
    OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'FR' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' 
    OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'RF' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' 
    OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'RH' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' 
    OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'HT' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    """

    sql3 = """Select EIRIN.ContNo,ContainerDetails.size,ContainerDetails.type,contCondition,
    cleaningtype.Name,Payload,If(Net=0,Null,Net),If(length(Datemnf) < 3,' ',Datemnf),EIRIN.DATEOUTPORT,
    EIRIN.DateIn,Concat(Exvessel,'-',ExVoy),UPPER(interchange.Consignee),
    EIRIN.VN,EIRIN.IntNo,EIRIN.PrincipleCode,EIRIN.Grade,EIRIN.Remark 
    From EIRIN 
    left Join cleaningtype On cleaningtype.PrincipleCode = EIRIN.PrincipleCode 
    left Join Interchange On Interchange.Nomor = EIRIN.IntNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = EIRIN.IntNo AND 
    InterchangeContainer.ContNo = EIRIN.ContNo 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.contno = eirin.contno 
    AND InterchangeDocpaycontainer.intno=EIRIN.intNo 
    left Join interchangeDocpaydetails on InterchangeDocpaydetails.nomor = InterchangeDocpaycontainer.nomor 
    AND InterchangeDocpaydetails.Size = ContainerDetails.Size 
    where EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""' 
    GROUP BY EIRIN.NOMOR order by EIRIN.DateIn"""

    sql3_summary = """
    SELECT 
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'GP' 
    AND EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'HC' 
    AND EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'OT' 
    AND EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '2021-07-12 00:00:01' AND EIRIN.DateIn <= '2021-07-14 23:59:59'
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'FR' 
    AND EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'RF' 
    AND EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'TK' 
    AND EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'GP' 
    AND EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'HC' 
    AND EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'OT' 
    AND EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '2021-07-12 00:00:01' AND EIRIN.DateIn <= '2021-07-14 23:59:59'
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'FR' 
    AND EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'RF' 
    AND EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'RH' 
    AND EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIRIN.contCondition LIKE '%AV%' OR EIRIN.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIRIN.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'HT' 
    AND EIRIN.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= '"""+time_from+"""' AND EIRIN.DateIn <= '"""+time_now+"""' 
    """

    sql4 = """Select EIROUT.ContNo as 'Container No',ContainerDetails.Size,ContainerDetails.Type,
    EIRIN.DateIn as 'Date In',Payload,ContainerDetails.Net as 'Tare',ContainerDetails.Datemnf as 'DMF',
    EIROUT.Nomor as 'EIR Out',Concat(ExVessel,'-',ExVoy) as 'Ex Vessel-Voy',BookingNo as 'DO No.',DateOut as 'Date Out',
    Concat(Vessel,'-',Voy) as 'Vessel-Voy',Destination,Shipper,EIROUT.VN as 'Truck No',
    EIROUT.ContCondition as 'Condition',EIROUT.Seal as 'Seal No.',EIROUT.PrincipleCode as Principal,EIROUT.Remark as 'Remarks' 
    From EIROUT 
    Left Join ContainerDetails on ContainerDetails.ContNo = EIROUT.ContNo 
    Left Join Booking On Booking.Nomor = EIROUT.BookingNo 
    Left Join EIRIN on EIRIN.Nomor = EIROUT.EIRIN 
    Left Join Interchange on Interchange.Nomor = EIRIN.IntNo 
    where EIROUT.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""' 
    group by EIROUT.ContNo,BookingNo order by EIROUT.DateOut"""
    
    sql4_summary = """
    SELECT 
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END)
    +COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIROUT 
    Left Join EIRIN on EIRIN.Nomor = EIROut.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'GP' 
    AND EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END)
    +COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIROUT 
    left Join EIRIN On EIRIN.Nomor = EIROut.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'HC' 
    AND EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND EIROUT.DateOut >= '"""+time_from+"""' 
    AND EIROUT.DateOut <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END)
    +COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIROUT 
    left Join EIRIN On EIRIN.Nomor = EIROut.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'OT' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END)+COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIROUT 
    left Join EIRIN On EIRIN.Nomor = EIROut.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'FR' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END)+COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIROUT 
    left Join EIRIN On EIRIN.Nomor = EIROut.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'RF' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END)+COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIROUT 
    left Join EIRIN On EIRIN.Nomor = EIROut.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'TK' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END)+COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIROUT 
    left Join EIRIN On EIRIN.Nomor = EIROut.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'GP' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END)+COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIROUT 
    left Join EIRIN On EIRIN.Nomor = EIROut.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'HC' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END)+COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIROUT 
    left Join EIRIN On EIRIN.Nomor = EIROut.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'OT' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END)+COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIROUT 
    left Join EIRIN On EIRIN.Nomor = EIROut.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'FR' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END)+COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIROUT 
    left Join EIRIN On EIRIN.Nomor = EIROut.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'RF' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END)+COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIROUT 
    left Join EIRIN On EIRIN.Nomor = EIROut.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'RH' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""' 
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN EIROUT.contCondition LIKE '%AV%' THEN 1 END)+COUNT(CASE WHEN EIROUT.contCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM EIROUT 
    left Join EIRIN On EIRIN.Nomor = EIROut.EIRIN 
    Left join ContainerDetails On ContainerDetails.ContNo = EIROUT.ContNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'HT' AND EIROUT.PrincipleCode = '"""+Principle_Code+"""' AND EIROUT.DateOut >= '"""+time_from+"""' AND EIROUT.DateOut <= '"""+time_now+"""' 
    """

    sql5 = """Select 
    ContainerStock.ContNo as 'Container No',ContainerDetails.Size as 'Size',ContainerDetails.Type as 'Type',
    ContainerStock.ContCondition as 'Condition',ContainerDetails.Payload as 'Payload',
    If(ContainerDetails.Net=0,Null,Net) as 'Tare',If(length(ContainerDetails.Datemnf) < 3,' ',
    ContainerDetails.Datemnf) as 'Date Mnf',Concat(Interchange.Exvessel,'-',Interchange.ExVoy) as 'Ex Vessel Voy',
    UPPER(Interchange.Consignee) as 'Customer',EIRIN.DateIn as 'Date IN',ContainerStock.PrincipleCode as Principle,
    BlockContainer.Remark as 'Remarks IN',EIRIN.Grade as Grade,EIRIN.IntNo as 'B/L NO' 
    From ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join BlockContainer On BlockContainer.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left Join Interchange On Interchange.Nomor = ContainerStock.IntNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = ContainerStock.IntNo AND 
    InterchangeContainer.ContNo = ContainerStock.ContNo 
    left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.ContNo = ContainerStock.ContNo 
    AND InterchangeDocpaycontainer.IntNo = ContainerStock.IntNo 
    left Join interchangeDocpaydetails on InterchangeDocpaydetails.Nomor = InterchangeDocpaycontainer.Nomor 
    AND InterchangeDocpaydetails.Size = ContainerDetails.Size 
    where ContainerStock.PrincipleCode = '"""+Principle_Code+"""' AND ContainerStock.ContCondition = 'DMG' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    GROUP BY ContainerStock.ContNo order by EIRIN.DateIn"""

    sql5_summary = """
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'GP' AND ContainerStock.ContCondition = 'DMG' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'HC' AND ContainerStock.ContCondition = 'DMG' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'OT' AND ContainerStock.ContCondition = 'DMG' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'FR' AND ContainerStock.ContCondition = 'DMG' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'RF' AND ContainerStock.ContCondition = 'DMG' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '20' AND ContainerDetails.type = 'TK' AND ContainerStock.ContCondition = 'DMG' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'GP' AND ContainerStock.ContCondition = 'DMG' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'HC' AND ContainerStock.ContCondition = 'DMG' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'OT' AND ContainerStock.ContCondition = 'DMG' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'   
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'FR' AND ContainerStock.ContCondition = 'DMG' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'RF' AND ContainerStock.ContCondition = 'DMG' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'RH' AND ContainerStock.ContCondition = 'DMG' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    where ContainerDetails.size = '40' AND ContainerDetails.type = 'HT' AND ContainerStock.ContCondition = 'DMG' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    """

    sql6 = """Select 
    ContainerStock.ContNo as 'Container No',ContainerDetails.Size as 'Size',ContainerDetails.Type as 'Type',
    ContainerStock.ContCondition as 'Condition',ContainerDetails.Payload as 'Payload',
    If(ContainerDetails.Net=0,Null,Net) as 'Tare',If(length(ContainerDetails.Datemnf) < 3,' ',
    ContainerDetails.Datemnf) as 'Date Mnf',Concat(Interchange.Exvessel,'-',Interchange.ExVoy) as 'Ex Vessel Voy',
    UPPER(Interchange.Consignee) as 'Customer',EIRIN.DateIn as 'Date IN',
    RepairContainer.CompleteRepair as 'Completed Repair',ContainerStock.PrincipleCode as Principle,
    RepairContainer.Repaired,BlockContainer.Remark as 'Remarks IN',EIRIN.Grade as Grade,EIRIN.IntNo as 'B/L NO' 
    From ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join BlockContainer On BlockContainer.ContNo = ContainerStock.ContNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    left Join Interchange On Interchange.Nomor = ContainerStock.IntNo 
    left Join InterchangeContainer On InterchangeContainer.Nomor = ContainerStock.IntNo AND 
    InterchangeContainer.ContNo = ContainerStock.ContNo 
    left Join interchangeDocpaycontainer on InterchangeDocpaycontainer.ContNo = ContainerStock.ContNo 
    AND InterchangeDocpaycontainer.IntNo = ContainerStock.IntNo 
    left Join interchangeDocpaydetails on InterchangeDocpaydetails.Nomor = InterchangeDocpaycontainer.Nomor 
    AND InterchangeDocpaydetails.Size = ContainerDetails.Size 
    where ContainerStock.PrincipleCode = '"""+Principle_Code+"""' AND RepairContainer.Repaired = 'Yes' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    GROUP BY ContainerStock.ContNo order by EIRIN.DateIn"""

    sql6_summary = """
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    where ContainerDetails.size = '20' AND RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'GP' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    where ContainerDetails.size = '20' AND RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'HC' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    where ContainerDetails.size = '20' AND RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'OT' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    where ContainerDetails.size = '20' AND RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'FR' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    where ContainerDetails.size = '20' AND RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'RF' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    where ContainerDetails.size = '20' AND RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'TK' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    where ContainerDetails.size = '40' AND RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'GP' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    where ContainerDetails.size = '40' AND RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'HC' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    where ContainerDetails.size = '40' AND RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'OT' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'   
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    where ContainerDetails.size = '40' AND RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'FR' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    where ContainerDetails.size = '40' AND RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'RF' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    where ContainerDetails.size = '40' AND RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'RH' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    UNION ALL 
    SELECT 
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' THEN 1 END) AS AV,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS DMG,
    COUNT(CASE WHEN ContainerStock.ContCondition LIKE '%AV%' OR ContainerStock.ContCondition LIKE '%DMG%' THEN 1 END) AS TOTAL 
    FROM ContainerStock 
    Left join ContainerDetails On ContainerDetails.ContNo = ContainerStock.ContNo 
    Left Join EIRIN On EIRIN.Nomor = ContainerStock.EIRNo 
    Left Join RepairContainer On RepairContainer.EORNo = ContainerStock.EORNo 
    where ContainerDetails.size = '40' AND RepairContainer.Repaired = 'Yes' AND ContainerDetails.type = 'HT' 
    AND ContainerStock.PrincipleCode = '"""+Principle_Code+"""' 
    AND EIRIN.DateIn >= (SELECT MIN(EIRIN.DateIn) FROM EIRIN) AND EIRIN.DateIn <= '"""+time_now+"""'  
    """
    file_name = f'REPORT-{Client}-{time_now_filename}.xlsx'
    # Writing Database into Excel
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    
    # Return sql query result
    #df1 = pd.read_sql_query(sql1, db)

    
    # Formatting Excel File
    workbook = writer.book

    # GLOBAL FORMATTING
    column_number_format = workbook.add_format({
        'bold': False,
        'align': 'center',
    })

    merge_info = workbook.add_format({
        'bold': True
    })

    merge_info_center = workbook.add_format({
        'bold': True,
    })

    merge_info_center.set_align('center')

    merge_info_center.set_align('vcenter')

    merge_format_number = workbook.add_format({
        'align': 'center',
        'bold': False
    })

    align_center = workbook.add_format({
        'align': 'center'
    })

    font_size_title = workbook.add_format({
        'font_size':20,
        'bold':True,
        'align':'center',
    })
    font_size_title.set_underline(1)

    # WORKSHEET 1 FORMAT
    #df1.index += 1
    #df1.to_excel(writer, sheet_name='DAILY STOCK POSITION', startrow=5, index=False)
    df1 = pd.read_sql_query(sql1, db) #Read the sql query
    df1.index += 1 #Add index by 1

    #df2.to_excel(writer, sheet_name='STOCK LIST', startrow=4)
    df1.to_excel(writer, sheet_name='DAILY STOCK POSITION', startrow=8, header=False, index=False, startcol=1)
    
    worksheet = writer.sheets['DAILY STOCK POSITION']
    worksheet.set_column(0, 0, 10)
    worksheet.set_row(0, 26.25)

    # Merge necessary column
    worksheet.merge_range('A1:S1', 'STOCK POSITION DAILY REPORT', font_size_title)
    worksheet.merge_range('A3:E3', f'PRINCIPAL: {Principle_Code}', merge_info)
    worksheet.merge_range('A4:E4', f'{time_now}', merge_info)
    worksheet.merge_range('A5:A8', 'TYPE CNTR', merge_info_center)
    worksheet.merge_range('B5:D5', 'STOCK', merge_info_center)
    worksheet.merge_range('B6:D6', 'AWAL', merge_info_center)
    worksheet.merge_range('B7:D7', '', merge_info_center)
    
    worksheet.write('B8:B8', "20'", merge_info_center)
    worksheet.write('C8:C8', "40'", merge_info_center)
    worksheet.write('D8:D8', "45'", merge_info_center)

    worksheet.merge_range('E5:V5', 'IN WARD', merge_info_center)
    worksheet.merge_range('W5:Y5', 'COMPLETED', merge_info_center)
    worksheet.merge_range('Z5:AB5', 'TOTAL', merge_info_center)
    worksheet.merge_range('AC5:AK6', 'STOCK AKHIR', merge_info_center)
    
    worksheet.merge_range('E6:G7', 'TOTAL IN', merge_info_center)#worksheet.merge_range('G6:H6', 'Ex - Vessel', merge_format)
    
    worksheet.merge_range('H6:M6', 'CONDITION', merge_info_center)
    worksheet.merge_range('N6:V6', 'CLEANING', merge_info_center)
    worksheet.merge_range('W6:Y6', 'REPAIR', merge_info_center)
    worksheet.merge_range('Z6:AB6', 'OUT WARD', merge_info_center)

    worksheet.merge_range('H7:J7', 'AV', merge_info_center)
    worksheet.merge_range('K7:M7', 'DM', merge_info_center)
    worksheet.merge_range('N7:P7', 'D/W & C/W', merge_info_center)
    worksheet.merge_range('Q7:S7', 'W/W', merge_info_center)
    worksheet.merge_range('T7:V7', 'S/W', merge_info_center)
    worksheet.merge_range('W7:Y7', '', merge_info_center)
    worksheet.merge_range('Z7:AB7', '', merge_info_center)
    worksheet.merge_range('AC7:AE7', 'TODAY', merge_info_center)
    worksheet.merge_range('AF7:AH7', 'AV', merge_info_center)
    worksheet.merge_range('AI7:AK7', 'DMG', merge_info_center)
    
    worksheet.write('E8:E8', "20'", merge_info_center)
    worksheet.write('F8:F8', "40'", merge_info_center)
    worksheet.write('G8:G8', "45'", merge_info_center)

    worksheet.write('H8:H8', "20'", merge_info_center)
    worksheet.write('I8:I8', "40'", merge_info_center)
    worksheet.write('J8:J8', "45'", merge_info_center)

    worksheet.write('K8:K8', "20'", merge_info_center)
    worksheet.write('L8:L8', "40'", merge_info_center)
    worksheet.write('M8:M8', "45'", merge_info_center)

    worksheet.write('N8:N8', "20'", merge_info_center)
    worksheet.write('O8:O8', "40'", merge_info_center)
    worksheet.write('P8:P8', "45'", merge_info_center)

    worksheet.write('Q8:Q8', "20'", merge_info_center)
    worksheet.write('R8:R8', "40'", merge_info_center)
    worksheet.write('S8:S8', "45'", merge_info_center)

    worksheet.write('T8:T8', "20'", merge_info_center)
    worksheet.write('U8:U8', "40'", merge_info_center)
    worksheet.write('V8:V8', "45'", merge_info_center)

    worksheet.write('W8:W8', "20'", merge_info_center)
    worksheet.write('X8:X8', "40'", merge_info_center)
    worksheet.write('Y8:Y8', "45'", merge_info_center)

    worksheet.write('Z8:Z8', "20'", merge_info_center)
    worksheet.write('AA8:AA8', "40'", merge_info_center)
    worksheet.write('AB8:AB8', "45'", merge_info_center)

    worksheet.write('AC8:AC8', "20'", merge_info_center)
    worksheet.write('AD8:AD8', "40'", merge_info_center)
    worksheet.write('AE8:AE8', "45'", merge_info_center)

    worksheet.write('AF8:AF8', "20'", merge_info_center)
    worksheet.write('AG8:AG8', "40'", merge_info_center)
    worksheet.write('AH8:AH8', "45'", merge_info_center)

    worksheet.write('AI8:AI8', "20'", merge_info_center)
    worksheet.write('AJ8:AJ8', "40'", merge_info_center)
    worksheet.write('AK8:AK8', "45'", merge_info_center)

    worksheet.write('A9:A9', "FR", merge_info)
    worksheet.write('A10:A10', "GP", merge_info)
    worksheet.write('A11:A11', "GOH", merge_info)
    worksheet.write('A12:A12', "HC", merge_info)
    worksheet.write('A13:A13', "OT", merge_info)
    worksheet.write('A14:A14', "RF", merge_info)
    worksheet.write('A15:A15', "RH", merge_info)
    worksheet.write('A16:A16', "TOTAL", None)

    worksheet.set_column('B:AK', None, align_center)

    '''
    dicts = ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'
    ,'AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK']
    for dict in dicts:
        worksheet.write(f"{dict}16", f"={dict}9+{dict}10+{dict}11+{dict}12+{dict}13+{dict}14+{dict}15",None)
    '''
    no_border_format = workbook.add_format({'bottom':0, 'top':0, 'left':0, 'right':0})
    worksheet.conditional_format( 'A1:AK4' , { 'type' : 'no_errors' , 'format' : no_border_format} )

    border_format = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
    worksheet.conditional_format( 'A5:AK16' , { 'type' : 'no_errors' , 'format' : border_format} )

    # SHEET 2 #
    df2 = pd.read_sql_query(sql2, db) #Read the sql query
    df2.index += 1 #Add index by 1

    #df2.to_excel(writer, sheet_name='STOCK LIST', startrow=4)
    df2.to_excel(writer, sheet_name='STOCK LIST', startrow=4)
    df2_summary = pd.read_sql_query(sql2_summary, db)
    index=["20GP","20HC","20OT","20FR","20RF","20TK","40GP","40HC","40OT","40FR","40RF","40RH","40HT"]
    df2_summary.index = index
    df2_summary.to_excel(writer, sheet_name='STOCK LIST', startrow=df2.shape[0] + 7,startcol=0)

    worksheet2 = writer.sheets['STOCK LIST']

    isempty = df2.empty
    if isempty!=False:
        worksheet2.set_row(0, 26.25)
        
        # Merge necessary column
        worksheet2.merge_range('A1:O1', 'CONTAINER STOCK LIST', font_size_title)
        worksheet2.merge_range('A3:E3', f'PRINCIPAL: {Principle_Code}', merge_info)
        worksheet2.merge_range('A4:E4', f'{time_now}', merge_info)

        # Add Border
        #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        #worksheet2.set_column(0,len(df2.columns),15,border_fmt)
        
        # Add Number Iteration
        worksheet2.write('A5:A5', 'No.', merge_format_number)

        worksheet2.set_column('A:O', 10, align_center)
        worksheet2.set_column(0, 0, 5)
        worksheet2.set_column(1, 1, 20)
        worksheet2.set_column(8, 8, 20)
        worksheet2.set_column(9, 9, 20)
        worksheet2.set_column(10, 10, 20)
        worksheet2.set_column(11, 11, 20)
        worksheet2.set_column(12, 12, 20)
        worksheet2.set_column(14, 14, 20)
    else:
            # WORKSHEET 2 FORMAT
        worksheet2.set_row(0, 26.25)
        
        # Merge necessary column
        worksheet2.merge_range('A1:L1', 'CONTAINER STOCK LIST', font_size_title)
        worksheet2.merge_range('A3:E3', f'PRINCIPAL: {Principle_Code}', merge_info)
        worksheet2.merge_range('A4:E4', f'{time_now}', merge_info)

        # Add Border
        #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        #worksheet2.set_column(0,len(df2.columns),15,border_fmt)
        
        # Add Number Iteration
        worksheet2.write('A5:A5', 'No.', merge_format_number)

        # Format column Text Align
        worksheet2.set_column('B:I', None, align_center) 
        worksheet2.set_column('K:L', None, align_center)
        worksheet2.set_column('N:O', None, align_center)
    
        # Auto-adjust column width
        for idx, col in enumerate(df2.columns):  # loop through all columns
            series = df2[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
                )) + 2  # adding a little extra space
            worksheet2.set_column(idx+1, idx+1, max_len)

        # Format column width
        worksheet2.set_column(0, 0, 5)
        worksheet2.set_column(6, 6, 10)
        worksheet2.set_column(7, 7, 10)
        worksheet2.set_column(10, 10, 20)


        # Un-bold the number iteration
        for row_num, value in enumerate(df2.index.get_level_values(level=0)):
            worksheet2.write(row_num+5, 0, value, column_number_format)

    # SHEET 3 #
    df3 = pd.read_sql_query(sql3, db)
    df3.index += 1 #Add index by 1

    header_mov_in = [
    "Container No","Size","Type","Condition","Cleaning","Payload","Tare","Date Mnf",
    "GateOut_CY","Date In","Ex Vessel Voy","Ex Consignee","Truck No",
    "B/L NO","Principal","Grade","Remarks"
    ]

    df3.to_excel(writer, sheet_name='MOV IN', startrow=4, header=header_mov_in)
    df3_summary = pd.read_sql_query(sql3_summary, db)
    index=["20GP","20HC","20OT","20FR","20RF","20TK","40GP","40HC","40OT","40FR","40RF","40RH","40HT"]
    df3_summary.index = index
    df3_summary.to_excel(writer, sheet_name='MOV IN', startrow=df3.shape[0] + 7,startcol=0)
    worksheet3 = writer.sheets['MOV IN']

    isempty = df3.empty
    if isempty!=False:
        # WORKSHEET 3 FORMAT
        worksheet3.set_row(0, 26.25)
        
        # Merge necessary column
        worksheet3.merge_range('A1:O1', 'MOVEMENT IN CONTAINER LIST', font_size_title)
        worksheet3.merge_range('A3:E3', f'PRINCIPAL: {Principle_Code}', merge_info)
        worksheet3.merge_range('A4:E4', f'{time_now}', merge_info)

        # Add Border
        #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        #worksheet2.set_column(0,len(df2.columns),15,border_fmt)
        
        # Add Number Iteration
        worksheet3.write('A5:A5', 'No.', merge_format_number)

        worksheet3.set_column('A:O', 10, align_center)
        worksheet3.set_column(0, 0, 5)
        worksheet3.set_column(1, 1, 20)
        worksheet3.set_column(8, 8, 20)
        worksheet3.set_column(9, 9, 20)
        worksheet3.set_column(10, 10, 20)
        worksheet3.set_column(11, 11, 20)
        worksheet3.set_column(12, 12, 20)
    else:
        worksheet3.set_row(0, 26.25)

        # Merge necessary column
        worksheet3.merge_range('A1:L1', 'MOVEMENT IN CONTAINER LIST', font_size_title)
        worksheet3.merge_range('A3:E3', f'PRINCIPAL: {Principle_Code}', merge_info)
        worksheet3.merge_range('A4:E4', f'{time_now}', merge_info)
        
        # Add Border
        #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        #worksheet3.set_column(0,len(df3.columns),15,border_fmt)

        # Add Number Iteration
        worksheet3.write('A5:A5', 'No.', merge_format_number)

        # Format column Text Align
        worksheet3.set_column('B:L', None, align_center) 
        worksheet3.set_column('N:Q', None, align_center)
        
        # Auto-adjust column width
        for idx, col in enumerate(df3.columns):  # loop through all columns
            series = df3[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
                )) + 2  # adding a little extra space
            worksheet3.set_column(idx+1, idx+1, max_len)

        # Format Column width
        worksheet3.set_column(0, 0, 5)
        worksheet3.set_column(7, 7, 10)
        worksheet3.set_column(8, 8, 10)
        worksheet3.set_column(10, 10, 20)

        # Un-bold the number iteration
        for row_num, value in enumerate(df3.index.get_level_values(level=0)):
            worksheet3.write(row_num+5, 0, value, column_number_format)
        
        # Add Border
        #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        #worksheet2.conditional_format(xlsxwriter.utility.xl_range(6, 0, len(df2.index), len(df2.columns)), {'type': 'no_errors', 'format': border_fmt})
        
    # SHEET 4 #
    df4 = pd.read_sql_query(sql4, db)
    df4.index += 1 #Add index by 1

    df4.to_excel(writer, sheet_name='MOV OUT', startrow=4)
    df4_summary = pd.read_sql_query(sql4_summary, db)
    index=["20GP","20HC","20OT","20FR","20RF","20TK","40GP","40HC","40OT","40FR","40RF","40RH","40HT"]
    df4_summary.index = index
    df4_summary.to_excel(writer, sheet_name='MOV OUT', startrow=df4.shape[0] + 7,startcol=0)
    worksheet4 = writer.sheets['MOV OUT']

    isempty = df4.empty
    if isempty!=False:
        # WORKSHEET 3 FORMAT
        worksheet4.set_row(0, 26.25)
        
        # Merge necessary column
        worksheet4.merge_range('A1:O1', 'MOVEMENT OUT CONTAINER LIST', font_size_title)
        worksheet4.merge_range('A3:E3', f'PRINCIPAL: {Principle_Code}', merge_info)
        worksheet4.merge_range('A4:E4', f'{time_now}', merge_info)

        # Add Border
        #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        #worksheet2.set_column(0,len(df2.columns),15,border_fmt)
        
        # Add Number Iteration
        worksheet4.write('A5:A5', 'No.', merge_format_number)

        worksheet4.set_column('A:O', 10, align_center)
        worksheet4.set_column(0, 0, 5)
        worksheet4.set_column(1, 1, 20)
        worksheet4.set_column(8, 8, 20)
        worksheet4.set_column(9, 9, 20)
        worksheet4.set_column(10, 10, 20)
        worksheet4.set_column(11, 11, 20)
        worksheet4.set_column(12, 12, 20)
        worksheet4.set_column(14, 14, 20)
    else:
        # WORKSHEET 4 FORMAT
        worksheet4.set_row(0, 26.25)

        # Merge necessary column
        worksheet4.merge_range('A1:L1', 'MOVEMENT OUT CONTAINER LIST', font_size_title)
        worksheet4.merge_range('A3:E3', f'PRINCIPAL: {Principle_Code}', merge_info)
        worksheet4.merge_range('A4:E4', f'{time_now}', merge_info)
        
        # Adding border to data
        #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        #worksheet2.set_column(0,len(df2.columns),5,border_fmt)

        # Add Number Iteration
        worksheet4.write('A5:A5', 'No.', merge_format_number)

        # Format column Text Align
        worksheet4.set_column('B:N', None, align_center) 
        
        # Auto-adjust column width
        for idx, col in enumerate(df4.columns):  # loop through all columns
            series = df4[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
                )) + 2  # adding a little extra space
            worksheet4.set_column(idx+1, idx+1, max_len)
        
        # Format Column width
        worksheet4.set_column(0, 0, 5)
        worksheet4.set_column(6, 6, 10)
        worksheet4.set_column(7, 7, 10)
        worksheet4.set_column(11, 11, 20)

        # Un-bold the number iteration
        for row_num, value in enumerate(df4.index.get_level_values(level=0)):
            worksheet4.write(row_num+5, 0, value, column_number_format)

    # SHEET 5 #
    df5 = pd.read_sql_query(sql5, db) #Read the sql query
    df5.index += 1 #Add index by 1

    #df2.to_excel(writer, sheet_name='STOCK LIST', startrow=4)
    df5.to_excel(writer, sheet_name='DAMAGE', startrow=4)
    df5_summary = pd.read_sql_query(sql5_summary, db)
    index=["20GP","20HC","20OT","20FR","20RF","20TK","40GP","40HC","40OT","40FR","40RF","40RH","40HT"]
    df5_summary.index = index
    df5_summary.to_excel(writer, sheet_name='DAMAGE', startrow=df5.shape[0] + 7,startcol=0)

    worksheet5 = writer.sheets['DAMAGE']

    isempty = df5.empty
    if isempty!=False:
        worksheet5.set_row(0, 26.25)
        
        # Merge necessary column
        worksheet5.merge_range('A1:O1', 'DAMAGE STOCK LIST', font_size_title)
        worksheet5.merge_range('A3:E3', f'PRINCIPAL: {Principle_Code}', merge_info)
        worksheet5.merge_range('A4:E4', f'{time_now}', merge_info)

        # Add Border
        #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        #worksheet2.set_column(0,len(df2.columns),15,border_fmt)
        
        # Add Number Iteration
        worksheet5.write('A5:A5', 'No.', merge_format_number)

        worksheet5.set_column('A:O', 10, align_center)
        worksheet5.set_column(0, 0, 5)
        worksheet5.set_column(1, 1, 20)
        worksheet5.set_column(8, 8, 20)
        worksheet5.set_column(9, 9, 20)
        worksheet5.set_column(10, 10, 20)
        worksheet5.set_column(11, 11, 20)
        worksheet5.set_column(12, 12, 20)
        worksheet5.set_column(14, 14, 20)
        
    else:
            # WORKSHEET 5 FORMAT
        worksheet5.set_row(0, 26.25)
        
        # Merge necessary column
        worksheet5.merge_range('A1:L1', 'DAMAGE STOCK LIST', font_size_title)
        worksheet5.merge_range('A3:E3', f'PRINCIPAL: {Principle_Code}', merge_info)
        worksheet5.merge_range('A4:E4', f'{time_now}', merge_info)

        # Add Border
        #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        #worksheet2.set_column(0,len(df2.columns),15,border_fmt)
        
        # Add Number Iteration
        worksheet5.write('A5:A5', 'No.', merge_format_number)

        # Format column Text Align
        worksheet5.set_column('B:I', None, align_center) 
        worksheet5.set_column('K:L', None, align_center)
        worksheet5.set_column('N:O', None, align_center)
    
        # Auto-adjust column width
        for idx, col in enumerate(df5.columns):  # loop through all columns
            series = df5[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
                )) + 2  # adding a little extra space
            worksheet5.set_column(idx+1, idx+1, max_len)

        # Format column width
        worksheet5.set_column(0, 0, 5)
        worksheet5.set_column(6, 6, 10)
        worksheet5.set_column(7, 7, 10)
        worksheet5.set_column(10, 10, 20)


        # Un-bold the number iteration
        for row_num, value in enumerate(df5.index.get_level_values(level=0)):
            worksheet5.write(row_num+5, 0, value, column_number_format)
    
    # SHEET 6 #
    df6 = pd.read_sql_query(sql6, db) #Read the sql query
    df6.index += 1 #Add index by 1

    #df2.to_excel(writer, sheet_name='STOCK LIST', startrow=4)
    df6.to_excel(writer, sheet_name='REPAIR FINISHED', startrow=4)
    df6_summary = pd.read_sql_query(sql6_summary, db)
    index=["20GP","20HC","20OT","20FR","20RF","20TK","40GP","40HC","40OT","40FR","40RF","40RH","40HT"]
    df6_summary.index = index
    df6_summary.to_excel(writer, sheet_name='REPAIR FINISHED', startrow=df6.shape[0] + 7,startcol=0)

    worksheet6 = writer.sheets['REPAIR FINISHED']

    isempty = df6.empty
    if isempty!=False:
        worksheet6.set_row(0, 26.25)
        
        # Merge necessary column
        worksheet6.merge_range('A1:O1', 'REPAIR FINISHED LIST', font_size_title)
        worksheet6.merge_range('A3:E3', f'PRINCIPAL: {Principle_Code}', merge_info)
        worksheet6.merge_range('A4:E4', f'{time_now}', merge_info)

        # Add Border
        #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        #worksheet2.set_column(0,len(df2.columns),15,border_fmt)
        
        # Add Number Iteration
        worksheet6.write('A5:A5', 'No.', merge_format_number)

        worksheet6.set_column('A:O', 10, align_center)
        worksheet6.set_column(0, 0, 5)
        worksheet6.set_column(1, 1, 20)
        worksheet6.set_column(8, 8, 20)
        worksheet6.set_column(10, 10, 20)
        worksheet6.set_column(11, 11, 20)
        worksheet6.set_column(12, 12, 20)
        worksheet6.set_column(14, 14, 20)
    else:
            # WORKSHEET 5 FORMAT
        worksheet6.set_row(0, 26.25)
        
        # Merge necessary column
        worksheet6.merge_range('A1:L1', 'REPAIR FINISHED LIST', font_size_title)
        worksheet6.merge_range('A3:E3', f'PRINCIPAL: {Principle_Code}', merge_info)
        worksheet6.merge_range('A4:E4', f'{time_now}', merge_info)

        # Add Border
        #border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        #worksheet2.set_column(0,len(df2.columns),15,border_fmt)
        
        # Add Number Iteration
        worksheet6.write('A5:A5', 'No.', merge_format_number)

        # Format column Text Align
        worksheet6.set_column('B:I', None, align_center) 
        worksheet6.set_column('K:L', None, align_center)
        worksheet6.set_column('N:O', None, align_center)
    
        # Auto-adjust column width
        for idx, col in enumerate(df6.columns):  # loop through all columns
            series = df6[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
                )) + 2  # adding a little extra space
            worksheet6.set_column(idx+1, idx+1, max_len)

        # Format column width
        worksheet6.set_column(0, 0, 5)
        worksheet6.set_column(6, 6, 10)
        worksheet6.set_column(7, 7, 10)
        worksheet6.set_column(10, 10, 10)

        # Un-bold the number iteration
        for row_num, value in enumerate(df6.index.get_level_values(level=0)):
            worksheet6.write(row_num+5, 0, value, column_number_format)

    # Save Excel File
    writer.save()

    fromaddr = 'jkt@autodkm.com'
    
    #toaddr = ['ivan.charlie@samudera.id', 'taufik.kurochman@samudera.id', 'yul.husein@samudera.id', 'jarukhi@samudera.id', 'sudarto@samudera.id', 'maulayas.shadak@samudera.id', 'cepih.balbarosah@samudera.id', 'arif.nursago@samudera.id', 'heri.sukiswanto@samudera.id', 'angga.sagita@samudera.id']
    #cc = ['agung-sil@dwipakharismamitra.co.id','jagoar_pasaribu@pt-dkm.co.id','richard@dwipakharismamitra.co.id','oktoberlin@dwipakharismamitra.co.id']
    #recipients = "oktoberlin@gmai.com"
    #fromaddr = 'report@autodkms.com'
    toaddr=['oktoberlin@gmail.com']
    cc=['linibelajar@gmail.com']
    #toaddr=to+cc
    msg = MIMEMultipart()

    msg['From'] = 'Report DKM Jakarta <jkt@autodkm.com>'
    msg['To'] = ", ".join(toaddr)
    msg['CC'] = ", ".join(cc)
    msg['Subject'] = f"DKM JAKARTA - REPORT OTOMATIS - {Client} - {time_now_email_subject}"

    #body = f"Auto Generate Excel Report PNPP DKM JAKARTA List {time_now}"


    #toaddr = ["oktoberlin@gmail.com","linibelajar@gmail.com"]
    #recipients = "oktoberlin@gmai.com"
    #msg = MIMEMultipart()

    #msg['From'] = 'Report DKM Jakarta <report@autodkms.com>'
    #msg['To'] = ", ".join(toaddr)
    #msg['Subject'] = f"DKM JAKARTA - REPORT OTOMATIS PER 10 MENIT - {Client} - {time_now_email_subject}"

    body = f"Auto Generate Excel Report {Client} DKM JAKARTA List {time_from} - {time_now}"

    msg.attach(MIMEText(body, 'plain'))

    filename = file_name
    attachment = open(file_name, "rb")

    part = MIMEBase('application','octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(part)
    try:
        server = smtplib.SMTP('srv42.niagahoster.com', 587)
        #server.set_debuglevel(3)
        server.starttls()
        server.login("jkt@autodkm.com", "jkt@dkm")
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr+cc, text)

        print('email successfully sent')
        server.quit()
    except SMTPResponseException as e:
        error_code = e.smtp_code
        error_message = e.smtp_error
        if (error_code==422):
            print("Recipient Mailbox Full")
        elif(error_code==431):
            print ("Server out of space")
        elif(error_code==447):
            print ("Timeout. Try reducing number of recipients")
        elif(error_code==510 or error_code==511):
            print ("One of the addresses in your TO, CC or BBC line doesn't exist. Check again your recipients' accounts and correct any possible misspelling.")
        elif(error_code==512):
            print ("Check again all your recipients' addresses: there will likely be an error in a domain name (like mail@domain.coom instead of mail@domain.com)")
        elif(error_code==541 or error_code==554):
            print ("Your message has been detected and labeled as spam. You must ask the recipient to whitelist you")
        elif(error_code==550):
            print ("Though it can be returned also by the recipient's firewall (or when the incoming server is down), the great majority of errors 550 simply tell that the recipient email address doesn't exist. You should contact the recipient otherwise and get the right address.")
        elif(error_code==553):
            print ("Check all the addresses in the TO, CC and BCC field. There should be an error or a misspelling somewhere.")
        else:
            print (error_code+": "+error_message)

if __name__ == '__main__':
    time_now_filename = datetime.now(pytz.timezone('Asia/Jakarta')).strftime("%Y%m%d")
    Client = 'PPNP'

    sent = False
    #day_temp = 0
    #day_now = datetime.now(pytz.timezone('Asia/Jakarta')).strftime("%d")
    #time_now = datetime.now(pytz.timezone('Asia/Jakarta')).strftime("%H:%M")
    
    time_now_email_subject = datetime.now(pytz.timezone('Asia/Jakarta')).strftime("%Y-%m-%d")
    #time_from = ''
    #time_from = datetime.now(pytz.timezone('Asia/Jakarta')).strftime("%Y-%m-%d")
    time_yesterday = (datetime.now(pytz.timezone('Asia/Jakarta'))-timedelta(1)).strftime("%Y-%m-%d %H:%M:%S")
    time_now = '2021-09-04 09:00:00'
    #if day_now!=day_temp:
    #    sent=False
    #    day_temp=day_now
    
    if f'{time_now_email_subject} 09:00:00' <= time_now <= f'{time_now_email_subject} 09:01:00' and sent==False:
        #global time_from
        time_from= time_now-timedelta(1)
        sent=True
        print('awesome')

        # Database name Query conditions
        mysql_to_excel(time_now_filename,Client,time_now_email_subject,time_from)   
'''

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

        /*
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
        */
        excel_filename = ''+ client +' - '+ start_date +' '+ end_date +' - REPORT.xlsx'
        # Writing Database into Excel
        writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
        # Return sql query result
        df1 = pd.read_sql_query(sql1, db)
        df2 = pd.read_sql_query(sql2, db)
        
        #df3 = pd.read_sql_query(sql3, db)
        #df4 = pd.read_sql_query(sql4, db)
        #df5 = pd.read_sql_query(sql5, db)
        #df6 = pd.read_sql_query(sql6, db)

        
        df1.index += 1
        df2.index += 1
        header = [
        "Container No","Size","Type","Condition","Payload","Tare","Date Mnf","Ex Vessel Voy",
        "Customer","Date In","Principal",
        "Remarks Restriction","Grade","B/L NO"
        ]

        df1.to_excel(writer, sheet_name='DAILY STOCK POSITION', startrow=5, header=header)
        df2.to_excel(writer, sheet_name='STOCK LIST', startrow=5)
        
        #df3.to_excel(writer, sheet_name='MOV IN', startrow=5)
        #df4.to_excel(writer, sheet_name='MOV OUT',startrow=5)
        #df5.to_excel(writer, sheet_name='DAMAGE', startrow=5)
        #df6.to_excel(writer, sheet_name='FINISHED REPAIR',startrow=5)
        
        worksheet2 = writer.sheets['STOCK LIST']
        
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

        
        #for column in df1:
        #    column_length = max(df1[column].astype(str).map(len).max(), len(column))
        #    col_idx = df1.columns.get_loc(column)
        #    worksheet2.set_column(col_idx+5, col_idx+5, column_length)
        
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
        
        fromaddr = 'report@autodkms.com'
        toaddr = email
        #recipients = "oktoberlin@gmai.com"
        msg = MIMEMultipart()

        msg['From'] = 'Report DKM Jakarta <report@autodkms.com>'
        msg['To'] = ", ".join(toaddr)
        msg['Subject'] = f"DKM JAKARTA - REPORT OTOMATIS PER 10 MENIT - {client} - {end_date}"

        body = f"Auto Generate Excel Report PNPP DKM JAKARTA List {start_date}-{end_date}"

        msg.attach(MIMEText(body, 'plain'))

        filename = excel_filename
        attachment = open(excel_filename, "rb")

        part = MIMEBase('application','octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        msg.attach(part)
        try:
            server = smtplib.SMTP('smtp.sendgrid.net', 587)
            #server.set_debuglevel(3)
            server.starttls()
            server.login("apikey", "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")
            text = msg.as_string()
            server.sendmail(fromaddr, toaddr, text)

            print('email successfully sent')
            server.quit()
        except SMTPResponseException as e:
            error_code = e.smtp_code
            error_message = e.smtp_error
            if (error_code==422):
                print("Recipient Mailbox Full")
            elif(error_code==431):
                print ("Server out of space")
            elif(error_code==447):
                print ("Timeout. Try reducing number of recipients")
            elif(error_code==510 or error_code==511):
                print ("One of the addresses in your TO, CC or BBC line doesn't exist. Check again your recipients' accounts and correct any possible misspelling.")
            elif(error_code==512):
                print ("Check again all your recipients' addresses: there will likely be an error in a domain name (like mail@domain.coom instead of mail@domain.com)")
            elif(error_code==541 or error_code==554):
                print ("Your message has been detected and labeled as spam. You must ask the recipient to whitelist you")
            elif(error_code==550):
                print ("Though it can be returned also by the recipient's firewall (or when the incoming server is down), the great majority of errors 550 simply tell that the recipient email address doesn't exist. You should contact the recipient otherwise and get the right address.")
            elif(error_code==553):
                print ("Check all the addresses in the TO, CC and BCC field. There should be an error or a misspelling somewhere.")
            else:
                print (error_code+": "+error_message)
        
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

'''