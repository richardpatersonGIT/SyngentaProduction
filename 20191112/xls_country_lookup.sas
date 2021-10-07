/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %IMPORT_COUNTRY_LOOKUP(...)*/
/*Purpose: Imports Country lookup mapping from excel file (cl_file=)*/
/*IN: Country lookup(1 excel file) sheet='Country_lookup'*/
/*OUT: dmimport.Country_lookup*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";

%macro IMPORT_COUNTRY_LOOKUP(cl_file=);

	PROC IMPORT OUT=Country_lookup_raw DATAFILE= "&cl_file." 
	            DBMS=xlsx REPLACE;
	     SHEET="Country_lookup"; 
	     GETNAMES=YES;
	RUN;

	data dmimport.Country_lookup;
		set Country_lookup_raw;
		if ^missing(unique_code) then output;
	run;

%mend IMPORT_COUNTRY_LOOKUP;