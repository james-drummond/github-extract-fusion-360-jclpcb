# Extract FUSION360 to jlcpcb.com

 This is a simple python program to produce the Bill of Materials (BOM) and layout from the FUSION 360 files which are needed to send a pcb assembly to jlcpcb.com
    
## Installation

Just use the program - it needs several modules
```python
  import os
  import zipfile
  import csv
  import openpyxl 
  import sys
  import io
  from argparse import ArgumentParser
  from time import strftime
```
  They are all available from 
```bash
  pip install <module>
```

## Usage

`python extract.py <zip> [-t through-hole] [-b previous bom] [-B]`


`zip:` \
The gerber zip from FUSION 360 (default "archive.zip")
You can this from FUSION 360 as:\
`"your PCB layout" > MANUFACTURING > MANUFACTURING > Export gerger....`
    
`-t through-hole:` \
If you only need the SMD parts, you don't need this file
To add the through-hole parts run:\
`"your PCB layout" > AUTOMATION > AUTOMATION > Run ULP` 
then run `mount_smd_tht.ulp` (use the defaults) which
produces the "tht" file needed 

`-b previous bom` \
tries allocated the jlcpcb part # from the original to the updating BOM
so you don't have to fill them in again 

`-B` \
uses the existing bom and verifies that it has not changed so the part numbers are unchanged. Useful when you are just tweaking the board layout and lettering, etc.
     


The finally result is the BOM and layout needed, but the BOM must be also populated with part numbers, usually from jlcpcb.com/parts.

The layout might need to be tweaked because the jlcpcb system doesn't always get the orientation correct.  The best thing is to upload the files to 
jclpcb.com and look at their version on the layout picture and then adjust the
layout file.

## Examples
`python extract.py gerber` \
Produces a new clean BOM and layout for the SMD parts

`python extract.py gerber -B` \
Creates the layout and will verify that the BOM is unchanged for the SMD parts

`python extract.py gerber -b old_bom` \
Creates the layout and try to allocate the jclpcb part numbers from the old bom

`python extract.py gerber -t through-hole` \
Produces a new clean BOM and layout for the SMD and the through-hole parts

## License

[gpl-3.0](https://choosealicense.com/licenses/gpl-3.0/)

