/***********************************************************************/
/*Type: Configuration*/
/*Use: Include in every program*/
/*Purpose: apply folder paths and sas libraries to programs*/
/***********************************************************************/

OPTION VALIDVARNAME=V7;

/*debug options - uncomment only 1*/
options mprint mlogic symbolgen;
/*options nomprint nomlogic nosymbolgen;*/

/*paths used in forecast report*/
%let dmfcst1_path=C:\Datamart\forecasts1;
%let dmfcst2_path=C:\Datamart\forecasts2;
%let dmfcst3_path=C:\Datamart\forecasts3;
%let dmfcst4_path=C:\Datamart\forecasts4;

/*libraries used in forecast report*/
libname dmfcst1 "&dmfcst1_path.";
libname dmfcst2 "&dmfcst2_path.";
libname dmfcst3 "&dmfcst3_path.";
libname dmfcst4 "&dmfcst4_path.";

/*libraries used in imports and reports*/
libname dmimport "C:\Datamart\imports";
libname dmproc "C:\Datamart\process";
libname SHW "C:\DATAMART\Sales_history_weekly";

/*path to metadata file(forecast report and uploads*/
%let metadata_file=C:\METADATA\metadata.xlsx;
%let or_metadata_file=C:\METADATA\orders_report.xlsx;
%let ex_metadata_file=C:\METADATA\extrapolation_report.xlsx;

%let metadata_folder=C:\METADATA;
%let import_folder=C:\IMPORT;
%let sas_applications_folder=C:\APPLICATIONS\SAS;
%let import_internal_folder=C:\IMPORT\internal;
%let sales_history_folder=C:\IMPORT\old\Sales_History_weekly_download; /*Monika excel weekly download files*/

%let upload_report_folder=C:\output\upload_report;
%let sales_report_folder=C:\output\sales_report;
%let forecast_report_folder=C:\output\forecast_report;
%let extrapolation_report_folder=C:\output\extrapolation_report;
%let error_report_folder=C:\output\errors_report;