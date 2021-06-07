/***********************************************************************/
/*Type: Utility*/
/*Use: Used as import_all program.*/
/*Purpose: Check the errors in import files.*/
/*IN: work.material_class_table_all*/
/*    work.variety_class_table_all*/
/*    work.material_class_table_dup*/
/*    work.variety_class_table_dup*/
/*    dmimport.FPS_Assortment*/
/*    work.FPS_Assortment1_dup*/
/*    DMPROC.PMD_assortment*/
/*    DMPROC.orders_all*/
/*OUT: dmproc.errors_all*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\cleanup_xlsx_bak_folder.sas";
%include "&sas_applications_folder.\filter_orders.sas";

%macro check_errors();

  proc datasets library=work nolist;
    delete  errs1
            errs2
            errs3
            errs4
            errs5
            errs6
            errs7
            errs8;
  run;

  data errs1(keep=ERROR_TYPE ERROR_SOURCE COLUMN_NAME COLUMN_VALUE_C COLUMN_VALUE_N ERROR_NUMBER ERROR_MESSAGE);
    length ERROR_TYPE $30. ERROR_SOURCE $30. COLUMN_NAME $32. COLUMN_VALUE_C $100. COLUMN_VALUE_N 8. ERROR_NUMBER  8. ERROR_MESSAGE $100.;
    set Material_class_table_all;
    ERROR_TYPE='IMPORT';
    ERROR_SOURCE='MATERIAL_CLASS_TABLE';
    if missing(sub_unit) then do;
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=1;
      ERROR_MESSAGE='No sub_unit for material';
      output;
    end;
    if missing(variety) then do;
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=2;
      ERROR_MESSAGE='No variety for material';
      output;
    end;
    if missing(material_desc) then do;
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=3;
      ERROR_MESSAGE='No Material description for material';
      output;
    end;
    if missing(division) then do;
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=4;
      ERROR_MESSAGE='No Mat_div for material';
      output;
    end;
    if missing(proc_stage) then do;
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=5;
      ERROR_MESSAGE='No Process Stage for material';
      output;
    end;
    if missing(proc_stage) then do;
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=6;
      ERROR_MESSAGE='No Process Stage for material';
      output;
    end;
  run;

  data errs2(keep=ERROR_TYPE ERROR_SOURCE COLUMN_NAME COLUMN_VALUE_C COLUMN_VALUE_N ERROR_NUMBER ERROR_MESSAGE);
    length ERROR_TYPE $30. ERROR_SOURCE $30. COLUMN_NAME $32. COLUMN_VALUE_C $100. COLUMN_VALUE_N 8. ERROR_NUMBER  8. ERROR_MESSAGE $100.;
    set Variety_class_table_all;
    ERROR_TYPE='IMPORT';
    ERROR_SOURCE='VARIETY_CLASS_TABLE';
    if missing(variety_desc) then do;
      COLUMN_NAME='VARIETY';
      COLUMN_VALUE_N=variety;
      ERROR_NUMBER=7;
      ERROR_MESSAGE='No variety name for variety';
      output;
    end;
    if missing(species_code) then do;
      COLUMN_NAME='VARIETY';
      COLUMN_VALUE_N=variety;
      ERROR_NUMBER=8;
      ERROR_MESSAGE='No species code for variety';
      output;
    end;
  run;

  %if %sysfunc(exist(material_class_table_dup)) %then %do;
    data errs3(keep=ERROR_TYPE ERROR_SOURCE COLUMN_NAME COLUMN_VALUE_C COLUMN_VALUE_N ERROR_NUMBER ERROR_MESSAGE);
      length ERROR_TYPE $30. ERROR_SOURCE $30. COLUMN_NAME $32. COLUMN_VALUE_C $100. COLUMN_VALUE_N 8. ERROR_NUMBER  8. ERROR_MESSAGE $100.;
      set material_class_table_dup;
      ERROR_TYPE='IMPORT';
      ERROR_SOURCE='MATERIAL_CLASS_TABLE';
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=9;
      ERROR_MESSAGE='Duplicated material';
    run;
  %end;

  %if %sysfunc(exist(variety_class_table_dup)) %then %do;
    data errs4(keep=ERROR_TYPE ERROR_SOURCE COLUMN_NAME COLUMN_VALUE_C COLUMN_VALUE_N ERROR_NUMBER ERROR_MESSAGE);
      length ERROR_TYPE $30. ERROR_SOURCE $30. COLUMN_NAME $32.COLUMN_VALUE_C $100. COLUMN_VALUE_N 8. ERROR_NUMBER  8. ERROR_MESSAGE $100.;
      set variety_class_table_dup;
      ERROR_TYPE='IMPORT';
      ERROR_SOURCE='VARIETY_CLASS_TABLE';
      COLUMN_NAME='VARIETY';
      COLUMN_VALUE_N=variety;
      ERROR_NUMBER=10;
      ERROR_MESSAGE='Duplicated variety';
    run;
  %end;

  data errs5(keep=ERROR_TYPE ERROR_SOURCE COLUMN_NAME COLUMN_VALUE_C COLUMN_VALUE_N ERROR_NUMBER ERROR_MESSAGE);
    length ERROR_TYPE $30. ERROR_SOURCE $30. COLUMN_NAME $32. COLUMN_VALUE_C $100. COLUMN_VALUE_N 8. ERROR_NUMBER  8. ERROR_MESSAGE $100.;
    set dmimport.FPS_Assortment;
    ERROR_TYPE='IMPORT';
    ERROR_SOURCE='FPS_ASSORTMENT';
    if missing(sub_unit) then do;
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=11;
      ERROR_MESSAGE='No sub_unit for material';
      output;
    end;
    if missing(variety) then do;
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=12;
      ERROR_MESSAGE='No Variety for material';
      output;
    end;
    if missing(curr_mat_plc) then do;
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=13;
      ERROR_MESSAGE='No PLC code for material';
      output;
    end;
    if missing(material_basic_description_en) then do;
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=14;
      ERROR_MESSAGE='No Material name for material';
      output;
    end;
    if missing(material_division) then do;
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=15;
      ERROR_MESSAGE='No Mat_div for material';
      output;
    end;
    if missing(dm_process_stage) then do;
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=16;
      ERROR_MESSAGE='No Process Stage for material';
      output;
    end;
    if missing(curr_mat_plc) then do;
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=17;
      ERROR_MESSAGE='No PLC for material';
      output;
    end;
  run;

  %if %sysfunc(exist(FPS_Assortment1_dup)) %then %do;
    data errs6(keep=ERROR_TYPE ERROR_SOURCE COLUMN_NAME COLUMN_VALUE_C COLUMN_VALUE_N ERROR_NUMBER ERROR_MESSAGE);
      length ERROR_TYPE $30. ERROR_SOURCE $30. COLUMN_NAME $32. COLUMN_VALUE_C $100. COLUMN_VALUE_N 8. ERROR_NUMBER  8. ERROR_MESSAGE $100.;
      set FPS_Assortment1_dup;
      ERROR_TYPE='IMPORT';
      ERROR_SOURCE='FPS_ASSORTMENT';
      COLUMN_NAME='MATERIAL';
      COLUMN_VALUE_N=material;
      ERROR_NUMBER=18;
      ERROR_MESSAGE='Duplicated material';
    run;
  %end;

  data errs7(keep=ERROR_TYPE ERROR_SOURCE COLUMN_NAME COLUMN_VALUE_C COLUMN_VALUE_N ERROR_NUMBER ERROR_MESSAGE);
    length ERROR_TYPE $30. ERROR_SOURCE $30. COLUMN_NAME $32. COLUMN_VALUE_C $100. COLUMN_VALUE_N 8. ERROR_NUMBER  8. ERROR_MESSAGE $100.;
    set DMPROC.PMD_assortment;
    ERROR_TYPE='IMPORT';
    ERROR_SOURCE='PMD_ASSORTMENT';
    if missing(product_line_group) and ^missing(product_line) then do;
      COLUMN_NAME='VARIETY';
      COLUMN_VALUE_N=variety;
      COLUMN_VALUE_C=variety_name;
      ERROR_NUMBER=19;
      ERROR_MESSAGE='Variety has product_line but dont have product_line_group mapping';
      output;
    end;
    if missing(product_line) then do;
      COLUMN_NAME='VARIETY';
      COLUMN_VALUE_N=variety;
      COLUMN_VALUE_C=variety_name;
      ERROR_NUMBER=20;
      ERROR_MESSAGE='No Product_line for variety';
      output;
    end;
    if missing(species) then do;
      COLUMN_NAME='VARIETY';
      COLUMN_VALUE_N=variety;
      COLUMN_VALUE_C=variety_name;
      ERROR_NUMBER=21;
      ERROR_MESSAGE='No Species name for variety';
      output;
    end;
    if missing(series) then do;
      COLUMN_NAME='VARIETY';
      COLUMN_VALUE_N=variety;
      COLUMN_VALUE_C=variety_name;
      ERROR_NUMBER=22;
      ERROR_MESSAGE='No Series name for variety';
      output;
    end;
    if missing(current_plc) and variety_name^='-' then do;
      COLUMN_NAME='VARIETY';
      COLUMN_VALUE_N=variety;
      COLUMN_VALUE_C=variety_name;
      ERROR_NUMBER=23;
      ERROR_MESSAGE='No PLC for variety';
      output;
    end;
  run;

  %filter_orders(in_table=dmproc.orders_all, out_table=orders_filtered);

  data errs8(keep=ERROR_TYPE ERROR_SOURCE COLUMN_NAME COLUMN_VALUE_C COLUMN_VALUE_N ERROR_NUMBER ERROR_MESSAGE);
    length ERROR_TYPE $30. ERROR_SOURCE $30. COLUMN_NAME $32. COLUMN_VALUE_C $100. COLUMN_VALUE_N 8. ERROR_NUMBER  8. ERROR_MESSAGE $100.;
    set orders_filtered;
    ERROR_TYPE='IMPORT';
    ERROR_SOURCE='ORDERS';
    if missing(country) then do;
      COLUMN_NAME='UNIQUE_CODE';
      COLUMN_VALUE_C=unique_code;
      ERROR_NUMBER=24;
      ERROR_MESSAGE='No mapping for unique_code in country_lookup';
      output;
    end;
    if region='FPS' and missing(sub_unit) then do;
      COLUMN_NAME='MATERIAL (descr, nr)';
      COLUMN_VALUE_N=material;
      COLUMN_VALUE_C=matdescr;
      ERROR_NUMBER=25;
      ERROR_MESSAGE='Material not in fps_assortment';
      output;
    end;
    if region^='FPS' and missing(sub_unit) then do;
      COLUMN_NAME='MATERIAL (descr, nr)';
      COLUMN_VALUE_N=material;
      COLUMN_VALUE_C=matdescr;
      ERROR_NUMBER=26;
      ERROR_MESSAGE='Material not in material_class_table';
      output;
    end;
    if missing(product_line_group) then do;
      COLUMN_NAME='VARIETY (region, nr)';
      COLUMN_VALUE_N=variety;
      COLUMN_VALUE_C=region;
      ERROR_NUMBER=27;
      ERROR_MESSAGE='Variety not in PMD, or region is mapped incorrectly';
      output;
    end;
  run;

  data errors_all;
    set 
        %if  %sysfunc(exist(errs1)) %then errs1;
        %if  %sysfunc(exist(errs2)) %then errs2;
        %if  %sysfunc(exist(errs3)) %then errs3;
        %if  %sysfunc(exist(errs4)) %then errs4;
        %if  %sysfunc(exist(errs5)) %then errs5;
        %if  %sysfunc(exist(errs6)) %then errs6;
        %if  %sysfunc(exist(errs7)) %then errs7;
        %if  %sysfunc(exist(errs8)) %then errs8;
      ;
  run;

  proc sql;
    create table dmproc.errors_all as
    select ERROR_TYPE, ERROR_SOURCE, COLUMN_NAME, COLUMN_VALUE_C, COLUMN_VALUE_N, ERROR_NUMBER, ERROR_MESSAGE, count(*) as ERROR_OCCURENCE
    from errors_all
    group by ERROR_TYPE, ERROR_SOURCE, COLUMN_NAME, COLUMN_VALUE_C, COLUMN_VALUE_N, ERROR_NUMBER, ERROR_MESSAGE
    order by ERROR_NUMBER, COLUMN_VALUE_N, COLUMN_VALUE_C;
  quit;

  data _null_;
    errors_filename=catx('_', compress(put(today(),yymmdd10.),,'kd'), compress(put(time(), time8.),,'kd'));
    call symput('errors_filename', strip(errors_filename));
  run;

    proc export 
    data=dmproc.errors_all
    dbms=xlsx 
    outfile="&error_report_folder.\errors_&errors_filename." replace ;
    sheet="Errors";
  run;

  %cleanup_xlsx_bak_folder(cleanup_folder=%str(&error_report_folder.\));

%mend check_errors;