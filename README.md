# alyson_xls_parser
Reads and Writes .xls files


Download the files:
Two ways to download
Using zip `Code > Download Zip`
Using git
```
git clone https://github.com/Palt0n/alyson_xls_parser.git
```

Install python libraries
```
pip install xlrd
pip install xlwt
```

```
pip install -r requirements.txt
```

To generate .xls file
```
python run_combine_xls.py example/data
```
To generate .xls file with name `out.xls`
```
python run_combine_xls.py example/data -o out.xls
```

## Setup
### Workspace Creation 
Navigate to new workspace folder
```
cd path/to/workspace
```

### Download from Git
```
git clone https://github.com/Palt0n/kh-serverstats.git
```

### Setup python virtual environment with venv
Create python virtual enviroment with venv
```
python -m venv env
```
To activate venv
```
source env/Scripts/activate
```
To check if venv is activated
```
which python
```
- This command should return the path for python.
- This path should be located in the local env folder

### Download external python libraries
Ensure that venv is activated, then install the libararies using pip
```
pip install -r requirements.txt
```
Or run 
```
pip install numpy
pip install matplotlib
```