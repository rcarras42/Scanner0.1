# Scanner0.1
 
#### Build

### Deps

- `pip install pytwain`
- `pip install pillow`
- `pip install -Iv cx_freeze=4.3.3`

### Env
## Debug, Tests

- `virtualenv -p python env`
- `.\env\Scripts\activate`
- `python pyscan.py`

## Build
This app has been built with CxFreeze
(Powershell, cwd: {path}\src)
- clear build dir `rm -Recurse ..\build\exe.win32-2.7`
- `python setup.py build`
