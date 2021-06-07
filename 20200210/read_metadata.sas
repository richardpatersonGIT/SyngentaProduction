/***********************************************************************/
/*Type: Utility*/
/*Use: Macro inside a program*/
/*Purpose: Create a sas dataset (dmimport.&sheet._md) based on excel sheet from metadata.xls file*/
/***********************************************************************/
%include "C:\APPLICATIONS\SAS\configuration.sas";

%macro read_metadata(sheet=);

  PROC IMPORT OUT=&sheet._md_raw 
              DATAFILE="&metadata_file."
              DBMS=  EXCELCS  REPLACE;
              SHEET="&sheet."; 
  RUN;

  data dmimport.&sheet._md;
    set &sheet._md_raw;
  run;

%mend read_metadata;