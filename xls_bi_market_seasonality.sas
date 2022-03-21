/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %IMPORT_BI_FIXED_SPLIT(...)*/
/*Purpose: Imports Bi Fixed Split from folder (bfs_folder=) with .xls* extension*/
/*IN: Variety_class_table(1 excel file with 2 sheets), extension=xls*, sheet="BI Market seasonality" and sheet="BI Market process stage split"*/
/*OUT:*/
/*  dmimport.BI_seasonality*/
/*  dmimport.BI_process_stage_split*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro IMPORT_BI_FIXED_SPLIT(bfs_folder=);

  %read_folder(folder=%str(&bfs_folder.), filelist=bfs_files, ext_mask=xls);

  proc sql noprint;
  select path_file_name into :bfs_file trimmed from bfs_files where order=1;
  quit;

  PROC IMPORT OUT=BI_seasonality_raw
              DATAFILE="&bfs_file."
              DBMS=  EXCELCS   REPLACE;
              SHEET="BI Market seasonality";
  RUN;

  PROC IMPORT OUT=BI_process_stage_split_raw
              DATAFILE="&bfs_file."
              DBMS=  EXCELCS   REPLACE;
              SHEET="BI Market process stage split";
  RUN;

  data dmimport.BI_seasonality (keep=product_line species month month_percentage);
    set BI_seasonality_raw;
    length month month_percentage 8.;
    month=1;
    month_percentage=coalesce(january, 0);
    output;

    month=2;
    month_percentage=coalesce(february, 0);
    output;

    month=3;
    month_percentage=coalesce(march, 0);
    output;

    month=4;
    month_percentage=coalesce(april, 0);
    output;

    month=5;
    month_percentage=coalesce(may, 0);
    output;

    month=6;
    month_percentage=coalesce(june, 0);
    output;

    month=7;
    month_percentage=coalesce(july, 0);
    output;

    month=8;
    month_percentage=coalesce(august, 0);
    output;

    month=9;
    month_percentage=coalesce(september, 0);
    output;

    month=10;
    month_percentage=coalesce(october, 0);
    output;

    month=11;
    month_percentage=coalesce(november, 0);
    output;

    month=12;
    month_percentage=coalesce(december, 0);
    output;
  run;

  data dmimport.BI_process_stage_split(keep=product_line species series variety process_stage process_stage_percentage);
    set BI_process_stage_split_raw;
    length process_stage $3. process_stage_percentage 8.;
    process_stage='CGS';
    process_stage_percentage=coalesce(CGS, 0);
    output;

    process_stage='CTD';
    process_stage_percentage=coalesce(CTD, 0);
    output;

    process_stage='MCO';
    process_stage_percentage=coalesce(MCO, 0);
    output;

    process_stage='MPL';
    process_stage_percentage=coalesce(MPL, 0);
    output;

    process_stage='PEL';
    process_stage_percentage=coalesce(PEL, 0);
    output;

    process_stage='PFN';
    process_stage_percentage=coalesce(PFN, 0);
    output;

    process_stage='PGS';
    process_stage_percentage=coalesce(PGS, 0);
    output;

    process_stage='PMD';
    process_stage_percentage=coalesce(PMD, 0);
    output;

    process_stage='PNV';
    process_stage_percentage=coalesce(PNV, 0);
    output;

    process_stage='PSY';
    process_stage_percentage=coalesce(PSY, 0);
    output;

    process_stage='RD';
    process_stage_percentage=coalesce(RD, 0);
    output;
    
    process_stage='RDB';
    process_stage_percentage=coalesce(RDB, 0);
    output;
    
    process_stage='RDD';
    process_stage_percentage=coalesce(RDD, 0);
    output;
    
    process_stage='RDY';
    process_stage_percentage=coalesce(RDY, 0);
    output;
  run;

%mend IMPORT_BI_FIXED_SPLIT;