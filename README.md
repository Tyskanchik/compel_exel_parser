![python](https://img.shields.io/badge/python-3.11.3-blue)
![customtkinter](https://img.shields.io/badge/customtkinter-5.2.0-blue)
![python](https://img.shields.io/badge/mypy-1.4.1-blue)


# Excel parser with sds and altium BOM
This program helps the designer to fill in the missing fields of BOM files when transferring them for PCB assembly.

## Discription
When uploading boom files from altium and ordering components from compel, it becomes necessary to compare the elements that needed to be bought with those that were bought.

The program allows you to compare BOM files from altium and excel files from sds-compel specifications. After matching, columns are created in the BOM file with the items that were bought and their quantity. Compel excel creates a column with the number of used elements on the board, to control the number of purchased elements.

## Usage
Download necessary files (Altium BOM and Compel Excel) in "change files" folder.<br>
Run `run_excel_parse.bat`, and in GUI choose Altium BOM and Compel Excel. Push botton and and enjoy the result.

## Example
The program uses column numbers. Therefore, it is desirable to define the table structure in advance.
### Tables structure
<table>
  <caption>Altium BOM</caption>
  <thead>
    <tr>
      <th scope="col">Name</th>
      <th scope="col">Description</th>
      <th scope="col">Footprint</th>
      <th scope="col">Designator</th>
      <th scope="col">Quantity</th>
      <th scope="col"><b>Название/закуплено</b></th>
      <th scope="col"><b>Количество</b></th>
      <th scope="col"><b>Примечание</b></th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>The table body</td>
      <td>with two columns</td>
    </tr>
  </tbody>
</table>
