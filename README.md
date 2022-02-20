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

Download and Install the libraries `xlrd` and `xlwt`
```
pip3 install -r requirements.txt
```

## How to run
Unzip the downloaded file `alyson_xls_parser-main.zip`
Open a terminal
Navigate to the location of the unzipped folder using the terminal
```
cd path/to/alyson_xls_parser-main
```
To generate .xls file from all files located in `example/original/Tanks/Tank 1/7dph` with multiple `1` named `combined_7dph.xls`
```
python3 run_flip_xls.py "example/original/Tanks/Tank 1/7dph" -m 1 -o "combined_7dph.xls"
```
To generate .xls file from all files located in `example/original/Tanks/Tank 1/13hpf` with multiple `4` named `combined_13hpf.xls`
```
python3 run_flip_xls.py "example/original/Tanks/Tank 1/13hpf" -m 4 -o "combined_13hpf.xls"
```
To generate .xls file from all files located in `example/original/Tanks/Tank 1/7hph` with multiple `3` named `combined_7hph.xls`
```
python3 run_flip_xls.py "example/original/Tanks/Tank 1/7hph" -m 3 -o "combined_7hph.xls"
```
Expected output
```
User@DESKTOP-ITS12AU MINGW64 /c/Github/alyson_xls_parser (main)
$ env/Scripts/python run_flip_xls.py "example/original/Tanks/Tank 1/7hph" -m 3 -o "combined_7hph.xls"
Searching for all .xls files in C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph
Found C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\1.xls
Found C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\2.xls
Found C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\3.xls
Found C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\4.xls
Found C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\5.xls
Found C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\6.xls
Found C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\7.xls
Found C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\8.xls
Found C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\9.xls
Read C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\1.xls - 3 data
Read C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\2.xls - 3 data
Read C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\3.xls - 3 data
Read C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\4.xls - 3 data
Read C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\5.xls - 3 data
Read C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\6.xls - 3 data
Read C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\7.xls - 3 data
Read C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\8.xls - 3 data
Read C:\Github\alyson_xls_parser\example\original\Tanks\Tank 1\7hph\9.xls - 3 data
Expected header: ['Type', 'Name', 'ID', 'Point X', 'Point Y', 'Length (µm)', 'Angle', 'Area', 'Perimeter']
Expected sheet name: Measurement
Data: 9
Files: 9/9 with 0 skipped
Generated file: combined_7hph.xls
Complete
```

