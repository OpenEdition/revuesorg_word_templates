@echo off

md build\bundle

md build\bundle\modeles
xcopy /Y src\macros\*.dot build\bundle\modeles
xcopy /Y build\templates\*.dot build\bundle\modeles

md build\bundle\startup
xcopy /Y src\startup\*.dot build\bundle\startup

echo Les modeles ont ete copies dans build\bundle
pause