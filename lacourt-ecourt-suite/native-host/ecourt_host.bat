@echo off
rem LA Court E-Court Suite - native messaging host launcher.
rem Chrome launches this .bat and speaks the native-messaging protocol over
rem stdio. %~dp0 is this file's own folder, so the .py next to it is found
rem regardless of the working directory Chrome uses.
rem
rem If a console window flashing on each Export bothers you, change "python"
rem to "pythonw" below. Use plain "python" if the merge ever fails to trigger
rem (pythonw can, in rare setups, mishandle the stdio pipes).
python "%~dp0ecourt_host.py"
