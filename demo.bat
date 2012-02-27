@echo off
python output/manage.py startapp %2
python read_xl.py %1 %2
echo Almost done. Now you need to:
echo 0. "cd output"
echo 1. Add %2 to the INSTALLED_APPS in settings.py.
echo 2. Run "python manage.py syncdb".
echo 3. Run "python manage.py loaddata %2/converted.json"
