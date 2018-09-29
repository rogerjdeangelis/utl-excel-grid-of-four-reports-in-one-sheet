Excel grid of four reports in one sheet

Output
https://tinyurl.com/ybfam2ql
https://github.com/rogerjdeangelis/utl-excel-grid-of-four-reports-in-one-sheet/blob/master/utl-excel-grid-of-four-reports-in-one-sheet.xlsx

github
https://tinyurl.com/yasxfxwm
https://github.com/rogerjdeangelis/utl-excel-grid-of-four-reports-in-one-sheet

You can use IML/R or my macro to layout the four reports in one sheet.

In 9.4M2
I don't think this is possible with ods excel.
There appears to be a bug in ods excel 'ods_excel_does_not_always_honor_start_at--bug' below.
You also cannot put reports side by side in 9.4M2 SAS.

Problem
https://tinyurl.com/ycuow6ap
https://github.com/rogerjdeangelis/ods_excel_does_not_always_honor_start_at--bug


SAS Forum
https://tinyurl.com/y76oy5ww
https://communities.sas.com/t5/SAS-Programming/excel-specify-each-report-the-start-and-end-point/m-p/499886

Other usefull links
https://github.com/rogerjdeangelis?utf8=%E2%9C%93&tab=repositories&q=excel+in%3Aname&type=&language=
https://github.com/rogerjdeangelis/utl_adding_SAS_graphics_at_an_arbitrary_position_into_existing_excel_sheets
https://github.com/rogerjdeangelis/utl_side_by_side_excel_reports

You can use ODS excel to stack reports or put reports in separate sheets and
then have R assemble them in one sheet and delete the other sheets.



INPUT
=====

WORK.HAVE total obs=48

Obs   SPECIES      WEIGHT     HEIGHT     WIDTH

 1    White         270.0     8.3804    4.2476
 2    White         270.0     8.1454    4.2485
 3    White         306.0     8.7780    4.6816
 4    White         540.0    10.7440    6.5620
 5    White         800.0    11.7612    6.5736
 6    White        1000.0    12.3540    6.5250
...
 7    Parkk          55.0     6.8475    2.3265
 8    Parkk          60.0     6.5772    2.3142
 9    Parkk          90.0     7.4052    2.6730
10    Parkk         120.0     8.3922    2.9181
11    Parkk         150.0     8.8928    3.2928
18    Pike          200.0     5.5680    3.3756
...
19    Pike          300.0     5.7078    4.1580
20    Pike          300.0     5.9364    4.3844
21    Pike          300.0     6.2884    4.0198
22    Pike          430.0     7.2900    4.5765
23    Pike          345.0     6.3960    3.9770
24    Pike          456.0     7.2800    4.3225
...
35    Smelt           6.7     1.7388    1.0476
36    Smelt           7.5     1.9720    1.1600
37    Smelt           7.0     1.7284    1.1484
38    Smelt           9.7     2.1960    1.3800
39    Smelt           9.8     2.0832    1.2772
40    Smelt           8.7     1.9782    1.2852


EXAMPLE OUTPUT
--------------


  Start At D11                                  Start At J11

      +--------------------------------------+   +--------------------------------------+
      |     D   |    E    |     F   |    G   |   |     J   |    K    |     L   |    M   |
      +--------------------------------------+   +--------------------------------------+
  11  | SPECIES |  WEIGHT |  HEIGHT |  WIDTH |   | SPECIES |  WEIGHT |  HEIGHT |  WIDTH |
      +---------+---------+---------+--------+   +---------+---------+---------+--------+
  12  | White   |   370   |   7.38  |    69  |   | Parkk   |    70   |   6.38  |    79  |
      +---------+---------+---------+--------+   +---------+---------+---------+--------+
  13  | White   |   250   |   8.14  |    58  |   | Parkk   |    84   |   5.14  |    98  |
      +---------+---------+---------+--------+   +---------+---------+---------+--------+
  14  | White   |   366   |   9.77  |    52  |   | Parkk   |    36   |   7.77  |    12  |
      +---------+---------+---------+--------+   +---------+---------+---------+--------+
  ...

      +---------+---------+---------+--------+   +---------+---------+---------+--------+   * supressed header
  25  | Pike    |   571   |   1.34  |    39  |   | Smelt   |   106   |   1.38  |    39  |
      +---------+---------+---------+--------+   +---------+---------+---------+--------+
  26  | Pike    |   471   |   2.64  |    58  |   | Smelt   |   304   |   1.14  |    56  |
      +---------+---------+---------+--------+   +---------+---------+---------+--------+
  27  | Pike    |   716   |   3.97  |    32  |   | Smelt   |   630   |   2.77  |    72  |
      +---------+---------+---------+--------+   +---------+---------+---------+--------+


PROCESS
=======


%utl_submit_r64('
source("c:/Program Files/R/R-3.3.2/etc/Rprofile.site",echo=T);
library(haven);
library(XLConnect);
have<-read_sas("d:/sd1/have.sas7bdat");
have;
White<-have[have$SPECIES=="White",];
parkk <-have[have$SPECIES=="Parkk",];
pike<-have[have$SPECIES=="Pike",];
smelt<-have[have$SPECIES=="Smelt",];
smelt;
parkk;
pike;
White;
wb <- loadWorkbook("d:/xls/utl-excel-grid-of-four-reports-in-one-sheet.xlsx", create = TRUE);
createSheet(wb, name = "species");
writeWorksheet(wb, parkk , sheet = "species", startRow = 11,startCol = 10, header = TRUE);
writeWorksheet(wb, White, sheet = "species", startRow = 11,startCol = 4, header = TRUE);
writeWorksheet(wb, smelt, sheet = "species", startRow = 25,startCol = 10, header = FALSE);
writeWorksheet(wb, pike, sheet = "species", startRow = 25,startCol = 4, header = FALSE);
saveWorkbook(wb);
');


OUTPUT
======

see above

*                _               _       _
 _ __ ___   __ _| | _____     __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \   / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/  | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|   \__,_|\__,_|\__\__,_|

;

options validvarname=upcase;
libname sd1 "d:/sd1";
data sd1.have;
  length species $5;
  set sashelp.fish (keep=species weight height width
      where=(species in ('Parkki' ,'Pike' ,'Smelt' ,'Whitefish')));
run;quit;

