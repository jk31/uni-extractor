# extractor-comparer

#### How to use
* download and use `extractor_comparer.exe`

#### Info for developing
* for PyInstaller you need a python environment only with the required packages, otherwise the compiled file gets huge
* install packages from `requirements.txt`
* install `pip install https://github.com/pyinstaller/pyinstaller/archive/develop.zip` instead of the usual PyInstaller to avoid errors (may change in future)
* for windows compiling run `pyinstaller -F --noconsole -n extractor_comparer app.py`


#### TODO:
* find better pdf text extrator
* error handling
* add functionality for default folder paths
