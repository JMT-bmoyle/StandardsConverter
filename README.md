# Standards Converter

This is a guide for how to convert standards in a `.JXL` file. Please reach out to me at BMoyle@jmt.com or on Teams if you have any questions.

# Installation
Installation has been streamlined to a `.BAT` file. In the event that the batch file fails to install the requirements, below will also be the steps for manual installation.

### Simple install
![Image of Install.bat](https://github.com/JMT-bmoyle/StandardsConverter/blob/main/Images/Install.png?raw=true)
In the directory `Q:\AZURE_HVMD\00-Reality_Capture\Sandbox\bmoyle\STANDARDS_CONVERTER` double click on `Install.bat`. This will install [Python3.10](https://www.python.org/downloads/release/python-3107/), Upgrade [PIP](https://pypi.org/project/pip/), and install the required modules [openpyxl](https://pypi.org/project/openpyxl/) and [tqdm](https://pypi.org/project/tqdm/).

### Manual install
In the directory `Q:\AZURE_HVMD\00-Reality_Capture\Sandbox\bmoyle\STANDARDS_CONVERTER`: 
- open `python-3.10.7-amd64.exe`
- Uncheck "Install launcher for all users (recommended)" and check "Add Python 3.10 to PATH"
- Click Install Now
- In commandline, run the following:
    - `python.exe -m pip install --upgrade pip`
    - `pip install openpyxl`
    - `pip install tqdm`
- Type `py` in commandline to ensure python is installed
- In the python terminal run `import openpyxl; import tqdm` to ensure the modules are installed.

# Usage
![Image of Converter.py](https://github.com/JMT-bmoyle/StandardsConverter/blob/main/Images/Converter.png?raw=true)

Double click on `Converter.py`. A window will pop up, then an explorer dialog will appear.

![Image of Standards Dialog](https://github.com/JMT-bmoyle/StandardsConverter/blob/main/Images/FindStandard.png?raw=true)

As seen in the bottom right corner, it is asking for our conversion standard. Navigate to the "Standards" folder.

![Image of selecting the desired standard](https://github.com/JMT-bmoyle/StandardsConverter/blob/main/Images/1SelectStandard.png?raw=true)

Open the desired standard. If one does not exist for your standard, follow the format in the excel file in order to make your own.

![Image of selecting the desired standard](https://github.com/JMT-bmoyle/StandardsConverter/blob/main/Images/FindJXL.png?raw=true)

As seen in the bottom right corner, it is asking for our `.JXL` file now. Navigate to your `JXL` file, and select it.

![Image of selecting the desired standard](https://github.com/JMT-bmoyle/StandardsConverter/blob/main/Images/Finished.png?raw=true)

The conversion process is finished! A pop up should tell you that it has completed. Click `OK`, and the file with `_Converted.jxl` as the suffex will be in the same location as the original. This is the file you want to use.

# Cleanup

![Image of selecting the desired standard](https://github.com/JMT-bmoyle/StandardsConverter/blob/main/Images/log.png?raw=true)

In the same location as the original and converted `JXL`'s, there will be a `.LOG` file. This file will contain any missing codes in your standard.

Any missing code can be fixed either manually in the `_Converted.jxl` file, or added to the Standard Excel file. If you choose to do the latter, re-run the converter with the origional again, not the converted file.

#### Note: Any numbers (eg. just the number "1" in your log) or ST codes can be ignored.
