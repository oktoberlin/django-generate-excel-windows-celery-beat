:: Create variables containing the name and path of the Celery pid file and the
:: names and paths to the batch files used to run Celery and Flower.
SET celeryPidFile=C:\xampp\htdocs\test_excel2\core\default.pid
SET celeryStartFile=C:\xampp\htdocs\test_excel2\core\celery_start.bat
SET flowerStartFile=D:\Server_Apps\flower\flower_start.bat

:: Celery 3.1.25 cannot be shutdown gracefully, and has to be killed. The
:: following command will kill all celery.exe processes and any python.exe
:: processes associated with Celery. The celery.exe process that is running
:: Flower will also be killed by this command.
TASKKILL /IM celery.exe /T /F

:: Force the deletion of the Celery pid file so that Celery can be restarted.
DEL /F "%celeryPidFile%"

:: Start Flower again.
START /B CMD /C CALL "%flowerStartFile%"

:: Start Celery again.
START /B CMD /C CALL "%celeryStartFile%"