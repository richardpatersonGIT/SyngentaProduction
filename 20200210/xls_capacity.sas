/***********************************************************************/
/*Type: Import*/
/*Use: Used in forecast_report program (macro call %read_capacity_vertical(...), %read_capacity_horizontal(...)*/
/*Purpose: Imports Capacity (vertical_capacity_file= or horizontal_capacity_file=)*/
/*IN: Capacity (1 vertical or 1 horizontal excel file) sheet=(first sheet in excel file, including hidden sheets)*/
/*OUT: dmimport.capacity*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";

%macro read_capacity_vertical(vertical_capacity_file=);

  PROC IMPORT OUT=tranposed_capacity_raw 
              DATAFILE="&vertical_capacity_file."
              DBMS=  EXCELCS   REPLACE;
  RUN;

  data capacity(drop=_variety week);
    length variety capacity_year capacity_week 8.;
    set tranposed_capacity_raw(rename=(variety=_variety));
    variety=input(_variety, 20.);
    capacity_year=input(substr(week, 1, 4),4.);
    capacity_week=input(substr(week, 5, 2),2.);
  run;

  data dmimport.capacity;
    set capacity;
  run;

%mend read_capacity_vertical;

%macro read_capacity_horizontal(horizontal_capacity_file=);
  PROC IMPORT OUT=capacity_raw 
              DATAFILE="&horizontal_capacity_file."
              DBMS=  EXCELCS   REPLACE;
  RUN;

  proc transpose data=capacity_raw out=capacity_transposed(drop=_name_ rename=(_label_=week col1=capacity));
    by variety;
  run;

  options VARLENCHK = NOWARN;
  data capacity_transposed1;
    length week $6.;
    set capacity_transposed;
    if length(week)=5 then do;
      week=cats(week,'0');
    end;
  run;
  options VARLENCHK = WARN;

  %let vertical_capacity_tmp=&import_internal_folder.\capacity_vertical_test.xlsx;

  proc export 
    data=capacity_transposed1
    dbms=xlsx 
    outfile="&vertical_capacity_tmp." replace;
    sheet='capacity_transposed_by_sas';
  run;

  %read_capacity_vertical(vertical_capacity_file=%quote(&vertical_capacity_tmp.));

%mend read_capacity_horizontal;


