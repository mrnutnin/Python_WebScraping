@echo off
set apphome=%~dp0
set app=Python27\python.exe
set appPy=src/controller.py
%apphome%%app% %appPy%
pause