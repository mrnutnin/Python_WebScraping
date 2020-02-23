@echo off
set apphome=%~dp0
set app=Python37\python.exe
set appPy=controller.py
%apphome%%app% %appPy%
pause