/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %IMPORT_COUNTRY_LOOKUP(...)*/
/*Purpose: Imports Country lookup mapping from excel file (cl_file=)*/
/*IN: Country lookup(1 excel file) sheet='Country_lookup'*/
/*IN: Country lookup(1 excel file) sheet='Soldto_nr_lookup'*/
/*OUT: dmimport.Country_lookup*/
/*OUT: dmimport.Soldto_nr_lookup*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";

%macro IMPORT_COUNTRY_LOOKUP(cl_file=);

  PROC IMPORT OUT=Country_lookup_raw DATAFILE= "&cl_file." 
              DBMS=xlsx REPLACE;
       SHEET="Country_lookup"; 
       GETNAMES=YES;
  RUN;

  data dmimport.Country_lookup;
    length region $3. territory $3. country $6;
    set Country_lookup_raw;
    if ^missing(unique_code) then output;
  run;

  PROC IMPORT OUT=soldto_nr_lookup_raw DATAFILE= "&cl_file." 
              DBMS=xlsx REPLACE;
       SHEET="soldto_nr_lookup"; 
       GETNAMES=YES;
  RUN;

  data dmimport.soldto_nr_lookup;
    length region $3. territory $3. country $6 Soldto_nr $8.; 
    set soldto_nr_lookup_raw;
    if ^missing(soldto_nr) then output;
  run;

%mend IMPORT_COUNTRY_LOOKUP;