/***********************************************************************/
/*Type: Report*/
/*Use: Fill in parameters in exrapolation_report.xlsx and press run*/
/*Purpose: Create Orders report from Running Sales*/
/*OUT: Excel file to extrapolation_report_folder (check configuration.sas)*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\cleanup_xlsx_bak_folder.sas";
%include "&sas_applications_folder.\filter_orders.sas";
%include "&sas_applications_folder.\extrapolation_extraction.sas";

%macro read_extrapolation_metadata();

  PROC IMPORT OUT=extrapolation_report_md_raw
              DATAFILE="&ex_metadata_file."
              DBMS=  EXCELCS  REPLACE; 
  RUN;

  data extrapolation_report_md(drop=_:);
    set extrapolation_report_md_raw(keep= region Product_line_group Seasonality Mat_div Hist_season week rename=(week=_week));
    length week 8.;
    week=input(_week, best.);
    if ^missing(coalesceC(of _character_)) or ^missing(coalesce(of _numeric_)) then output;
  run;

%mend read_extrapolation_metadata;

%macro extrapolation_report();

  %let extrapolation_start_time=EXTRAPOLATION STARTED: %sysfunc(date(),worddate.). %sysfunc(time(),timeampm.);

  %read_extrapolation_metadata();

  %extrapolation_extraction(extrapolation_config_ds=extrapolation_report_md);

  proc sql noprint;
      select mat_div into :mat_div_name separated by '_' from extrapolation_report_md where ^missing(mat_div);
      select catx('_', region, product_line_group, "&mat_div_name.", seasonality, put(hist_season, 4.), cats("wk",put(coalesce(week, 0), z2.))) into :ext_report_name trimmed from extrapolation_report_md where ^missing(region);
  quit;

  data _null_;
    extrapolation_report_file=catx('_', compress(put(today(),yymmdd10.),,'kd'), compress(put(time(), time8.),,'kd'));
    call symput('extrapolation_report_file', strip(extrapolation_report_file));
  run;

  %let extrapolation_name=&extrapolation_report_folder.\Extrapolation_&ext_report_name._&extrapolation_report_file..xlsx;

  x "del &extrapolation_name."; 

  proc export 
    data=species_sales_percentage
    dbms=xlsx 
    outfile="&extrapolation_name." replace;
    sheet="Species";
  run;

  proc export 
    data=country_sales_percentage
    dbms=xlsx 
    outfile="&extrapolation_name.";
    sheet="All_countries";
  run;

  proc sql noprint;
    create table countries as
    select distinct country from all_country_weeks;
    select country into :countries separated by '#' from countries;
    select count(*) into: countries_cnt trimmed from countries;
  quit;

  %do ci=1 %to &countries_cnt.;
    %let country=%scan(&countries., &ci., '#');
    data country_report;
      set country_sales_percentage;
      if country="&country." then output;
    run;

    proc export 
      data=country_report
      dbms=xlsx 
      outfile="&extrapolation_name.";
      sheet="&country.";
    run;
  %end;

  proc export 
    data=extrapolation_report_md
    dbms=xlsx 
    outfile="&extrapolation_name.";
    sheet="Variant";
  run;

  %cleanup_xlsx_bak_folder(cleanup_folder=%str(&extrapolation_report_folder.\));

  %let extrapolation_end_time=EXTRAPOLATION ENDED: %sysfunc(date(),worddate.). %sysfunc(time(),timeampm.);

  %put &=extrapolation_report_folder.;
  %put &=extrapolation_start_time.;
  %put &=extrapolation_end_time.;

%mend extrapolation_report;

%extrapolation_report();



