/***********************************************************************/
/*Type: Import*/
/*Use: Used in forecast_report program (macro call %read_supply_vertical(...), %read_supply_horizontal(...)*/
/*Purpose: Imports Supply (vertical_supply_file= or horizontal_supply_file=)*/
/*IN: Supply (1 vertical or 1 horizontal excel file) sheet=(first sheet in excel file, including hidden sheets)*/
/*OUT: dmimport.supply*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";

%macro read_supply_vertical(vertical_supply_file=);

  PROC IMPORT OUT=tranposed_supply_raw 
              DATAFILE="&vertical_supply_file."
              DBMS=  EXCELCS   REPLACE;
  RUN;

  data supply(drop=_variety week);
    length variety supply_year supply_week 8.;
    set tranposed_supply_raw(rename=(variety=_variety));
    variety=input(_variety, 20.);
    supply_year=input(substr(week, 1, 4),4.);
    supply_week=input(substr(week, 5, 2),2.);
  run;

  data dmimport.supply;
    set supply;
  run;

%mend read_supply_vertical;

%macro read_supply_horizontal(horizontal_supply_file=);
  PROC IMPORT OUT=supply_raw 
              DATAFILE="&horizontal_supply_file."
              DBMS=  EXCELCS   REPLACE;
  RUN;

  proc transpose data=supply_raw out=supply_transposed(drop=_name_ rename=(_label_=week col1=supply));
    by variety;
  run;

  options VARLENCHK = NOWARN;
  data supply_transposed1;
    length week $6.;
    set supply_transposed;
    if length(week)=5 then do;
      week=cats(week,'0');
    end;
  run;
  options VARLENCHK = WARN;

  %let vertical_supply_tmp=&import_internal_folder.\supply_vertical_test.xlsx;

  proc export 
    data=supply_transposed1
    dbms=xlsx 
    outfile="&vertical_supply_tmp." replace;
    sheet='supply_transposed_by_sas';
  run;

  %read_supply_vertical(vertical_supply_file=%quote(&vertical_supply_tmp.));

%mend read_supply_horizontal;