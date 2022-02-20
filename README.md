# alyson_xls_parser
Combines multiple small `.xls` files into a big `.xls` file

## How to setup
To setup this script, this requires you to:
- Download this code to your PC
- Install `python 3` 
- Setup and activate your `python 3` workspace
- Install `python 3` libraries `xlrd` and `xlwt`.

### Download Script
To download this code into your PC, github has a option for you to download as zip `Code > Download Zip`

### Install Python 3
- `Python 3` downloaded and installed from https://www.python.org/

### Setup and Activate workspace
Navigate to your unzipped folder
```
cd path/to/workspace
```
Create the virutal env workspace
```
python -m venv env
```
Activate the workspace
```
source env/Scripts/activate
```
### Install Python 3 Libraries
After installing `Python 3` and activating your workspace, you need to download the libraries

Download and Install the libraries `xlrd` and `xlwt`
```
pip install -r requirements.txt
```


## How to run
Navigate to the location of the script
```
cd path/to/workspace
```
To generate .xls file from all files located in `example/data`
```
python3 run_combine_xls.py example/data
```
To generate .xls file with name `out.xls`
```
python3run_combine_xls.py example/data -o out.xls
```
