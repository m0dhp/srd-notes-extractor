# SRD Notes Extract


## Purpose

Searches the SRD spreadsheet to extract the complete set of note numbers referenced in all routes
and a list of note numbers of all defined notes.


## Installation

Install Python 3.11 or above.

Install `pip3`.

Download and extract the repository into a directory.

Change to the directory.  Create and activate a virtual environment:

```
$ python3 -m venv venv
$ source venv/bin/activate
```

Install the required Python packages: 

```
pip3 install -r requirements.txt
```


## Usage

Save the SRD file in `.xlsx` format and copy it into the program folder.

Edit the program to set the SRD file name.  For example:
```
SRD_FILE_NAME = "SRD-Spreadsheet-2023-07-13_CRC_9B2DAC22.xlsx"
```

Run the program:

```
python3 extract.py
```
The program writes three text files into the program folder:

| File name                              | Contents |
|----------------------------------------|---|
| `notes-SRD-Spreadsheet ...`       | List of just the note numbers from the "Notes" tab |
| `notes-md-SRD-Spreadsheet ...`    | Same list but each item is preceded by markdown for a tick box list |
| `references-SRD-Spreadsheet ...`  | Sorted list of note numbers which are referenced one or more times in the "Routes" tab |

For testing, set the debug flag in the program.  If `DEBUG = True`, the program output
is sent to the console rather than written to file.


## History



##### 0.1 (2023-08-09)

First working version.



## License

This work is licensed by Ray Benitez under the Creative Commons Attribution-ShareAlike 4.0 International License. To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/4.0/ or send a letter to Creative Commons, PO Box 1866, Mountain View, CA 94042, USA.

