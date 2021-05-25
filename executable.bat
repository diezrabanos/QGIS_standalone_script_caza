REM Change OSGEO4W_ROOT to point to the base install folder


SET QGIS=C:\Program Files\QGIS 3.10\apps\qgis-ltr

REM Gdal Setup

set GDAL_DATA=C:\Program Files\QGIS 3.10share\gdal\

REM Python Setup

set PATH=C:\Program Files\QGIS 3.10\apps\qgis-ltr\bin;%PATH%
SET PYTHONHOME=C:\Program Files\QGIS 3.10\apps\Python37
set PYTHONPATH=C:\Program Files\QGIS 3.10\apps\qgis-ltr\python;%PYTHONPATH%

REM Launch python job

"C:\Program Files\QGIS 3.10\apps\Python37\python.exe" O:\sigmena\utilidad\PROGRAMA\QGIS\otros\standalone\caza_licencias\alone_qgis_script_licencia_caza_mup.py
pause
