:: Create variables containing the drive and path to OpenREM and the name and
:: path of the Celery pid and log files.
SET openremDrive=D:
SET openremPath=D:\Server_Apps\python27\Lib\site-packages\openrem
SET celeryPidFile=E:\media_root\celery\default.pid
SET celeryLogFile=E:\media_root\celery\default.log

:: Change to the drive on which OpenREM is installed and navigate to the
:: OpenREM folder.
%openremDrive%
CD "%openremPath%"

:: Start Celery.
celery worker -n default -Ofair -A openremproject -c 4 -Q default --pidfile=%celeryPidFile% --logfile=%celeryLogFile%