/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %IMPORT_ZDEMAND_LOOKUP(...)*/
/*Purpose: Imports Zdemand lookup mapping from excel file (cl_file=)*/
/*IN: Zdemand lookup(1 excel file) sheet='Zdemand_lookup'*/
/*OUT: dmimport.Zdemand_lookup*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";

%macro IMPORT_ZDEMAND_LOOKUP(cl_file=);

  PROC IMPORT OUT=zdemand_lookup_raw DATAFILE= "&cl_file." 
              DBMS=xlsx REPLACE;
       SHEET="zDemand_lookup"; 
       GETNAMES=YES;
  RUN;

  data dmimport.zdemand_lookup;
    length 
      Info_Str $4.
      Div $2.
      Sales_Org $4.
      Sales_Off $4.
      Region $3.
      Mat_div $2.
    ;
    set zdemand_lookup_raw;
    if ^missing(info_str) then output;
  run;

%mend IMPORT_ZDEMAND_LOOKUP;