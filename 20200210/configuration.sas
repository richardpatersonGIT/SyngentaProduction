/***********************************************************************/
/*Type: Configuration*/
/*Use: Include in every program*/
/*Purpose: apply folder paths and sas libraries to programs*/
/***********************************************************************/

%macro configuration();
OPTION VALIDVARNAME=V7;

/*debug options - uncomment only 1*/
options mprint mlogic symbolgen;
/*options nomprint nomlogic nosymbolgen;*/

/*paths used in forecast report*/
%let dmfcst1_path=C:\Datamart\forecasts1;
%let dmfcst2_path=C:\Datamart\forecasts2;
%let dmfcst3_path=C:\Datamart\forecasts3;
%let dmfcst4_path=C:\Datamart\forecasts4;

/*SAS libraries used in forecast report*/
libname dmfcst1 "&dmfcst1_path.";
libname dmfcst2 "&dmfcst2_path.";
libname dmfcst3 "&dmfcst3_path.";
libname dmfcst4 "&dmfcst4_path.";

/*SAS libraries used in imports and reports*/
libname dmimport "C:\Datamart\imports";
libname dmproc "C:\Datamart\process";
libname SHW "C:\Datamart\Sales_history_weekly";

/*path to metadata file(forecast report and uploads*/
/*%let metadata_file=S:\METADATA\metadata.xlsx;*/
/*%let or_metadata_file=S:\METADATA\orders_report.xlsx;*/
/*%let ex_metadata_file=S:\METADATA\extrapolation_report.xlsx;*/
/*%let sd_metadata_file=S:\METADATA\SvsD_report.xlsx;*/

/*Paths to metadata, import and report folders*/
/*%let metadata_folder=S:\METADATA;*/
/*%let import_folder=S:\IMPORT;*/
/*%let sas_applications_folder=S:\APPLICATIONS\SAS;*/
/*%let import_internal_folder=S:\IMPORT\internal;*/
/*%let sales_history_folder=S:\IMPORT\old\Sales_History_weekly_download; */
/*%let upload_report_folder=S:\output\upload_report;*/
/*%let sales_report_folder=S:\output\sales_report;*/
/*%let forecast_report_folder=S:\output\forecast_report;*/
/*%let extrapolation_report_folder=S:\output\extrapolation_report;*/
/*%let error_report_folder=S:\output\errors_report;*/
/*%let SvsD_report_folder=S:\output\SvsD_report;*/


/*LOCAL DRIVE*/



%let metadata_file=D:\METADATA\metadata.xlsx;
%let or_metadata_file=D:\METADATA\orders_report.xlsx;
%let ex_metadata_file=D:\METADATA\extrapolation_report.xlsx;
%let sd_metadata_file=D:\METADATA\SvsD_report.xlsx;


%let metadata_folder=D:\METADATA;
%let import_folder=D:\IMPORT;
%let sas_applications_folder=D:\APPLICATIONS\SAS;
%let import_internal_folder=D:\IMPORT\internal;
%let sales_history_folder=D:\IMPORT\old\Sales_History_weekly_download;
%let upload_report_folder=D:\output\upload_report;
%let sales_report_folder=D:\output\sales_report;
%let forecast_report_folder=D:\output\forecast_report;
%let extrapolation_report_folder=D:\output\extrapolation_report;
%let error_report_folder=D:\output\errors_report;
%let SvsD_report_folder=D:\output\SvsD_report;

%mend configuration();

%configuration;