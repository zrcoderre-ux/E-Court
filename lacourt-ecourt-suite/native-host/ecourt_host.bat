@echo off
rem LA Court E-Court Suite - native messaging host launcher.
rem Chrome launches this .bat and speaks the native-messaging protocol over
rem stdio. %~dp0 is this file's own folder, so the .py next to it is found
rem regardless of the working directory Chrome uses.
rem
rem The interpreter is a hardcoded ABSOLUTE path on purpose: this machine uses
rem Microsoft Store Python, whose bare "python" app-execution alias can fail to
rem resolve when Chrome (not an interactive shell) spawns this host. If you move
rem to a different Python, update the path below to that python.exe. To find it,
rem run:  python -c "import sys; print(sys.executable)"
rem
rem Console flash on each Export? Swap python.exe -> pythonw.exe below (same
rem folder). Revert to python.exe if the merge ever stops triggering, since
rem pythonw can, in rare setups, mishandle the stdio pipes.
"C:\Users\ZCoderre\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\python.exe" "%~dp0ecourt_host.py"
