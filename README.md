# Excel parser with sds and altium BOM
This program helps the designer to fill in the missing fields of BOM files when transferring them for PCB assembly.

## Discription
When uploading boom files from altium and ordering components from compel, it becomes necessary to compare the elements that needed to be bought with those that were bought.

The program allows you to compare BOM files from altium and excel files from sds-compel specifications. After matching, columns are created in the BOM file with the items that were bought and their quantity. Compel excel creates a column with the number of used elements on the board, to control the number of purchased elements.

## Usage
Run `run_excel_parse.bat`, and in GUI choose Altium BOM and Compel Excel. Push botton and and enjoy the result.
