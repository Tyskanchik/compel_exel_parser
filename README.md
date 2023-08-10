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
      <th scope="col"><i>Название/закуплено</i></th>
      <th scope="col"><i>Количество</i></th>
      <th scope="col"><i>Примечание</i></th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Piezoelectric Buzzer</td>
      <td>PKMCS0909E4000-R1</td>
      <td>PKMCS0909E4000-R1</td>
      <td>BQ1</td>
      <td>1</td>
      <td>PKMCS0909E4000-R1</td>
      <td>25</td>
      <td></td>
    </tr>
  </tbody>
    <tbody>
    <tr>
      <td>470Ohm</td>
      <td>RC1206FR-07470RL</td>
      <td>1206 RESISTOR</td>
      <td>R1, R2, R3</td>
      <td>3</td>
      <td>CR-06FL7--470R</td>
      <td>70</td>
      <td></td>
    </tr>
  </tbody>
    <tbody>
    <tr>
      <td>TL431</td>
      <td>TL431</td>
      <td>SOT-23-3</td>
      <td>DA1</td>
      <td>1</td>
      <td>RS431AYSF3</td>
      <td>25</td>
      <td></td>
    </tr>
  </tbody>
    </tbody>
    <tbody>
    <tr>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
  </tbody>
</table>
<br>
<table>
  <caption>SDS Excel</caption>
  <thead>
    <tr>
      <th scope="col">Имя строки</th>
      <th scope="col">Наименование</th>
      <th scope="col">Наименование товара</th>
      <th scope="col">Производитель товара</th>
      <th scope="col">Корпус</th>
      <th scope="col">Примечание</th>
      <th scope="col">Кол-во в изделии</th>
      <th scope="col">Подобрано предложение</th>
      <th scope="col">Приоритетный товар</th>
      <th scope="col">Общее кол-во на партию</th>
      <th scope="col">Используемое количество</th>
      <th scope="col"><i>Примечание</i></th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>RC1206FR-07470RL</td>
      <td>CR-06FL7--470R (VIKING)</td>
      <td>CR-06FL7--470R</td>
      <td>VIKING</td>
      <td>R1, R2, R3</td>
      <td>SMD120632X16MM</td>
      <td>Корпус: 1206 RESISTOR 470</td>
      <td>3</td>
      <td>да</td>
      <td>нет</td>
      <td>70</td>
      <td>60</td>
    </tr>
  </tbody>
    <tbody>
    <tr>
      <td>RC1206FR-07470RL</td>
      <td>AC1206FR-07470RL (YAG)</td>
      <td>AC1206FR-07470RL</td>
      <td>YAG</td>
      <td>R1, R2, R3</td>
      <td>SMD120632X16MM</td>
      <td>Корпус: 1206 RESISTOR 470</td>
      <td>3</td>
      <td>нет</td>
      <td>нет</td>
      <td>70</td>
      <td>60</td>
    </tr>
  </tbody>
    <tbody>
    <tr>
      <td>RC1206FR-07470RL</td>
      <td>CRCW1206470RFKEA (VISHAY)</td>
      <td>CRCW1206470RFKEA</td>
      <td>VISHAY</td>
      <td>R1, R2, R3</td>
      <td>SMD120632X16MM</td>
      <td>Корпус: 1206 RESISTOR 470</td>
      <td>3</td>
      <td>нет</td>
      <td>нет</td>
      <td>70</td>
      <td>60</td>
    </tr>
  </tbody>
    </tbody>
    <tbody>
    <tr>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
  </tbody>
</table>
*Italicized columns are added by the program
