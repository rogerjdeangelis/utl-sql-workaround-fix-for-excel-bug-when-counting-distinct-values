# utl-sql-workaround-fix-for-excel-bug-when-counting-distinct-values
SQL workaround fix for excel bug with missing values
    %let pgm=utl-sql-workaround-fix-for-excel-bug-when-counting-distinct-values;

    %stop_submission;

    SQL workaround fix for excel bug with missing values

    SOAPBOX ON

      Working with missing values is critical for most statistical analysis.
      Unfortunately an empty cell in excel is not a missing value.
      SAS, Sqlite, MySQL and postgreSQL do not count missing values with count distinct.

      For proper handling of missing values in excel, it is best to pass the workbook to SAS, Python, R or Sqlite?

    SOAPBOX OFF

    github
    https://tinyurl.com/ycxh299r
    https://stackoverflow.com/questions/79348328/excel-distinct-counts-on-pivot-tables-are-aggregating-incorrectly

    related github
    https://tinyurl.com/mr2b9f3y
    https://github.com/rogerjdeangelis/utl-dealing-with-missing-values-consitently-within-and-between-multiple-languages-sas-R-and-python

    /*                 _   _
      _____  _____ ___| | | |__  _   _  __ _
     / _ \ \/ / __/ _ \ | | `_ \| | | |/ _` |
    |  __/>  < (_|  __/ | | |_) | |_| | (_| |
     \___/_/\_\___\___|_| |_.__/ \__,_|\__, |
                                       |___/
    */

    /**************************************************************************************************************************/
    /*                                           |                          |                                                 */
    /*                                           |                          |                                                 */
    /*              INPUT (EXCEL SHEET)          |     PROCESS              |          OUTPUT (EXCEL SHEET)                   */
    /*              ===================          |     =======              |          ====================                   */
    /*                                           |                          |                                                 */
    /*                                           |                          |                                                 */
    /*  --------------------+                    |                          |                                                 */
    /*  | A1  | fx | NAME   |                    | =CALCULATE(DISTINCTCOUNT | NOTE: BLANK CELLS ARE COUNTED IN DISTINCT COUNT */
    /*  ---------------------------------------- |  (Table1[teacher])       |                                                 */
    /*  [_]|    A  |    B    |    C    |    E  | |                          |   --------------------+                         */
    /*  ---------------------------------------- |                          |   | A1  | fx | NAME   |                         */
    /*   1 | NAME  | STUDENT | TEACHER |  AIDE | |                          |   ------------------------------------------    */
    /*  -- |-------+---------+---------+-------+ |                          |   [_]|    A  |    B    |    C    |    E    |    */
    /*   2 | john  |         |         |       | |                          |   ------------------------------------------    */
    /*  -- |-------+---------+---------+-------+ |                          |    1 | NAME  | STUDENT | TEACHER |  AIDE T |    */
    /*   3 | john  |         |john     |       | |                          |   -- |-------+---------+---------+---------+    */
    /*  -- |-------+---------+---------+-------+ |                          |    2 |   4   |    1    |    2    |   2     |    */
    /*   4 | amy   |         |         |teacher| |                          |   -- |-------------------------------------+    */
    /*  -- |-------+---------+---------+-------+ |                          |   [WANT]                                        */
    /*   5 | jan   |         |         |       | |                          |                SHOULD BE                        */
    /*  -- |-------+---------+---------+-------+ |                          |                                                 */
    /*   6 | john  |         |john     |       | |                          |   --------------------+                         */
    /*  -- |-------+---------+---------+-------+ |                          |   | A1  | fx | NAME   |                         */
    /*   7 | mary  |         |         |       | |                          |   ------------------------------------------    */
    /*  -- |-------+---------+---------+-------+ |                          |   [_]|    A  |    B    |    C    |    E    |    */
    /*  [HAVE]                                   |                          |   ------------------------------------------    */
    /*                                           |                          |    1 | NAME  | STUDENT | TEACHER |  AIDE   |    */
    /*                                           |                          |   -- |-------+---------+---------+---------+    */
    /*                                           |                          |    2 |   4   |    0    |    1    |   1     |    */
    /*                                           |                          |   -- |-------------------------------------+    */
    /*                                           |                          |                                                 */
    /*                                           |                          |                                                 */
    /**************************************************************************************************************************/

    /*         _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __
    / __|/ _ \| | | | | __| |/ _ \| `_ \
    \__ \ (_) | | |_| | |_| | (_) | | | |
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|

    */

    /**************************************************************************************************************************/
    /*                                       |                                        |                                       */
    /*                 INPUT                 |                PROCESS                 |              OUTPUT                   */
    /*                 =====                 |                =======                 |              ======                   */
    /*                                       |                                        |                                       */
    /*  PREPARE THE EXCEL SHEET FOR SQL      |          SAME CODE SAS R PYTHO         |  WANT SHEET ADDED TO d:/xls/have.xlsx */
    /*  (SAS R OR PYTHON)                    |                                        |                                       */
    /*                                       |                                        |                                       */
    /*  d:/xls/have.xlsx sheet=have          |PASS WORKBOOK TO R AND CALL SQLDF       |                                       */
    /*                                       |                                        |                                       */
    /*  ----------------+                    | select                                 |  --------------------------+          */
    /*  | A1  | fx| NAME|                    |   count(distinct(student)) as student  |  | A1     | fx     |NAME   |          */
    /*  --------------------------------     |  ,count(distinct(teacher)) as teacher  |  --------------------------------     */
    /*  [_]|    A |   B  |   C |    E  |     |  ,count(distinct(aide))    as aide     |  [_]|   A |    B   |   C   | E  |     */
    /*  --------------------------------     | from                                   |  --------------------------------     */
    /*   1 | NAME | AGE  | SEX | HEIGHT|     |   have                                 |   1 | NAME| STUDENT|TEACHER|AIDE|     */
    /*  -- |------+------+-----+-------+     |                                        |  -- |-----+--------+-------+----+     */
    /*   2 | john |NA    |NA   |NA     |     |                                        |   2 |   4 |    0   |   1   | 1  |     */
    /*  -- |------+------+-----+-------+     |                                        |  -- |---------------------------+     */
    /*   3 | john |NA    |john |NA     |     |                                        |  [WANT]                               */
    /*  -- |------+------+-----+-------+     |                                        |                                       */
    /*   4 | amy  |NA    |NA   |teacher|     |                                        |                                       */
    /*  -- |------+------+-----+-------+     |                                        |                                       */
    /*   5 | jan  |NA    |NA   |NA     |     |                                        |                                       */
    /*  -- |------+------+-----+-------+     |                                        |                                       */
    /*   6 | john |NA    |john |NA     |     |                                        |                                       */
    /*  -- |------+------+-----+-------+     |                                        |                                       */
    /*   7 | mary |NA    |NA   |NA     |     |                                        |                                       */
    /*  -- |------+------+-----+-------+     |                                        |                                       */
    /*  [HAVE]                               |                                        |                                       */
    /*                                       |                                        |                                       */
    /**************************************************************************************************************************/

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    options validvarname=upcase;
    libname sd1 "d:/sd1";
    data sd1.have;
      input name$ student$ teacher$ aide$;
    cards4;
    john NA NA NA NA
    john NA john NA NA
    amy NA NA teacher NA
    jan NA NA NA aid
    john NA john NA NA
    mary NA NA NA aid
    ;;;;
    run;quit;

    %utl_rbeginx;
    parmcards4;
    library(openxlsx)
    library(sqldf)
     wb<-loadWorkbook("d:/xls/wantxl.xlsx")
     have<-read.xlsx(wb,"have")
     have
     addWorksheet(wb, "want")
     want<-sqldf('
       select
          count(distinct(student)) as unq_student
         ,count(distinct(teacher)) as unq_teacher
         ,count(distinct(aide))    as unq_aide
       from
          have
      ')
     print(want)
     writeData(wb,sheet="want",x=want)
     saveWorkbook(
         wb
        ,"d:/xls/wantxl.xlsx"
        ,overwrite=TRUE)
    ;;;;
    %utl_rendx;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*                                                                                                                        */
    /*    NAME    STUDENT    TEACHER    AIDE                                                                                  */
    /*                                                                                                                        */
    /*    john      NA        NA        NA                                                                                    */
    /*    john      NA        john      NA                                                                                    */
    /*    amy       NA        NA        teacher                                                                               */
    /*    jan       NA        NA        NA                                                                                    */
    /*    john      NA        john      NA                                                                                    */
    /*    mary      NA        NA        NA                                                                                    */
    /*                                                                                                                        */
    /*  PASSED TO EXCEL                                                                                                       */
    /*                                                                                                                        */
    /*   ----------------+                                                                                                    */
    /*   | A1  | fx| NAME|                                                                                                    */
    /*   --------------------------------                                                                                     */
    /*   [_]|    A |   B  |   C |    E  |                                                                                     */
    /*   --------------------------------                                                                                     */
    /*    1 | NAME | AGE  | SEX | HEIGHT|                                                                                     */
    /*   -- |------+------+-----+-------+                                                                                     */
    /*    2 | john |NA    |NA   |NA     |                                                                                     */
    /*   -- |------+------+-----+-------+                                                                                     */
    /*    3 | john |NA    |john |NA     |                                                                                     */
    /*   -- |------+------+-----+-------+                                                                                     */
    /*    4 | amy  |NA    |NA   |teacher|                                                                                     */
    /*   -- |------+------+-----+-------+                                                                                     */
    /*    5 | jan  |NA    |NA   |NA     |                                                                                     */
    /*   -- |------+------+-----+-------+                                                                                     */
    /*    6 | john |NA    |john |NA     |                                                                                     */
    /*   -- |------+------+-----+-------+                                                                                     */
    /*    7 | mary |NA    |NA   |NA     |                                                                                     */
    /*   -- |------+------+-----+-------+                                                                                     */
    /*   [HAVE]                                                                                                               */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    %utl_rbeginx;
    parmcards4;
    library(openxlsx)
    library(sqldf)
     wb<-loadWorkbook("d:/xls/wantxl.xlsx")
     have<-read.xlsx(wb,"have")
     have
     addWorksheet(wb, "want")
     want<-sqldf('
       select
          count(distinct(student)) as unq_student
         ,count(distinct(teacher)) as unq_teacher
         ,count(distinct(aide))    as unq_aide
       from
          have
      ')
     print(want)
     writeData(wb,sheet="want",x=want)
     saveWorkbook(
         wb
        ,"d:/xls/wantxl.xlsx"
        ,overwrite=TRUE)
    ;;;;
    %utl_rendx;


    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*  THE OUTPUT CONTAINS TWO SHEETS                                                                                        */
    /*  NOTE: R UPDATES AN EXISTNG WORKBOOK                                                                                   */
    /*                                                                                                                        */
    /*                                         COUNT DISTINCT VALUES                                                          */
    /*                                                                                                                        */
    /*   ----------------+                     --------------------------+                                                    */
    /*   | A1  | fx| NAME|                     | A1     | fx     |NAME   |                                                    */
    /*   --------------------------------      --------------------------------                                               */
    /*   [_]|    A |   B  |   C |    E  |      [_]|   A |    B   |   C   | E  |                                               */
    /*   --------------------------------      --------------------------------                                               */
    /*    1 | NAME | AGE  | SEX | HEIGHT|       1 | NAME| STUDENT|TEACHER|AIDE|                                               */
    /*   -- |------+------+-----+-------+      -- |-----+--------+-------+----+                                               */
    /*    2 | john |NA    |NA   |NA     |       2 |   4 |    0   |   1   | 1  |                                               */
    /*   -- |------+------+-----+-------+      -- |---------------------------+                                               */
    /*    3 | john |NA    |john |NA     |      [WANT]                                                                         */
    /*   -- |------+------+-----+-------+                                                                                     */
    /*    4 | amy  |NA    |NA   |teacher|                                                                                     */
    /*   -- |------+------+-----+-------+                                                                                     */
    /*    5 | jan  |NA    |NA   |NA     |                                                                                     */
    /*   -- |------+------+-----+-------+                                                                                     */
    /*    6 | john |NA    |john |NA     |                                                                                     */
    /*   -- |------+------+-----+-------+                                                                                     */
    /*    7 | mary |NA    |NA   |NA     |                                                                                     */
    /*   -- |------+------+-----+-------+                                                                                     */
    /*   [HAVE]                                                                                                               */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
