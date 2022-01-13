/***********************************************************************/
/*Type: Configuration*/
/*Use: Include in every program*/
/*Purpose: apply folder paths and sas libraries to programs*/
/***********************************************************************/

OPTION VALIDVARNAME=V7;

/* test */
/*debug options - uncomment only 1*/
options mprint mlogic symbolgen compress=yes reuse=yes;
/*options nomprint nomlogic nosymbolgen compress=yes reuse=yes;*/

/*SAS libraries used in forecast report*/
%let datamart_path=C:\SAS\Datamart;
%if %sysfunc(libref(dmfcst1))^=0 %then %do;
  libname dmfcst1 "&datamart_path.\forecasts1";
%end;
%if %sysfunc(libref(dmfcst2))^=0 %then %do;
  libname dmfcst2 "&datamart_path.\forecasts2";
%end;
%if %sysfunc(libref(dmfcst3))^=0 %then %do;
  libname dmfcst3 "&datamart_path.\forecasts3";
%end;
%if %sysfunc(libref(dmfcst4))^=0 %then %do;
  libname dmfcst4 "&datamart_path.\forecasts4";
%end;

/*libname dmfcst1 "&datamart_path.\forecasts1";*/
/*libname dmfcst2 "&datamart_path.\forecasts2";*/
/*libname dmfcst3 "&datamart_path.\forecasts3";*/
/*libname dmfcst4 "&datamart_path.\forecasts4";*/

/*SAS libraries used in imports and reports*/
%if %sysfunc(libref(dmimport))^=0 %then %do;
  libname dmimport "&datamart_path.\imports";
%end;
%if %sysfunc(libref(dmproc))^=0 %then %do;
  libname dmproc "&datamart_path.\process";
%end;
%if %sysfunc(libref(SHW))^=0 %then %do;
  libname SHW "&datamart_path.\Sales_history_weekly";
%end;

/*libname dmimport "&datamart_path.\imports";*/
/*libname dmproc "&datamart_path.\process";*/
/*libname SHW "&datamart_path.\Sales_history_weekly";*/

%let import_folder=C:\SAS\IMPORT;
%let sas_applications_folder=C:\SAS\APPLICATIONS\SAS;
%let import_internal_folder=C:\SAS\IMPORT\internal;
%let sales_history_folder=C:\SAS\IMPORT\old\Sales_History_weekly_download;
%let upload_report_folder=C:\SAS\output\upload_report;
%let sales_report_folder=C:\SAS\output\sales_report;
%let forecast_report_folder=C:\SAS\output\forecast_report;
%let extrapolation_report_folder=C:\SAS\output\extrapolation_report;
%let error_report_folder=C:\SAS\output\errors_report;
%let SvsD_report_folder=C:\SAS\output\SvsD_report;
%let metadata_folder=C:\SAS\METADATA;

%let metadata_file=&metadata_folder.\metadata.xlsx;
%let or_metadata_file=&metadata_folder.\orders_report.xlsx;
%let ex_metadata_file=&metadata_folder.\extrapolation_report.xlsx;
%let sd_metadata_file=&metadata_folder.\SvsD_report.xlsx;
