/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %import_tactical_plan(...)*/
/*Purpose: Imports Tactical Plan from folder (tp_folder=) with .xlsx extension*/
/*IN: tactical_plan(1 file), extension=xlsx, sheet='Tactical Plan'*/
/*OUT: dmimport.tactical_plan*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro import_tactical_plan(tp_folder=);

  %read_folder(folder=%str(&tp_folder.), filelist=tp_files, ext_mask=xlsx);

  proc sql noprint;
  select path_file_name into :tp_file trimmed from tp_files where order=1;
  quit;

  PROC IMPORT OUT=tactical_plan_raw
              DATAFILE="&tp_file."
              DBMS=  EXCELCS   REPLACE;
              sheet='Tactical Plan';
  RUN;

  data tactical_plan(keep=region product_line species_code mat_div country _:);
    set tactical_plan_raw;
  run;

  proc sort data=tactical_plan out=tactical_plan_sorted;
    by region product_line species_code mat_div country;
  run;

  proc transpose data=tactical_plan_sorted out=tactical_plan_transposed;
    by region product_line species_code mat_div country;
  run;

  data tactical_plan1(drop=_tp_year);
    length tp_year 8. hash_mat_div $3.;
    set tactical_plan_transposed (drop=_label_ rename=(_name_=_tp_year col1=tp_value));
    tp_year=input(compress(_tp_year,,'kd'), 4.);
    hash_mat_div=upcase(compress(compress(mat_div, ' ', 'ka')));
  run;

  proc sql;
    create table tactical_plan2 as
    select 
      a.*, 
      put(a.tp_year, 4.)||'-'||put(b.tp_year, 4.) as tp_spread length=9,
      b.tp_value/a.tp_value-1 as tp_growth,
      "Strategic/Tactical" as crop_categories
    from tactical_plan1 a
    left join tactical_plan1 b 
    on a.region=b.region and
    a.product_line=b.product_line and
    a.species_code=b.species_code and
    a.hash_mat_div=b.hash_mat_div and
    a.country=b.country
    and 
      (a.tp_year=b.tp_year-1 or
      a.tp_year=b.tp_year-2)
    where ^missing(b.tp_year)
    order by a.region, a.product_line, a.species_code, a.hash_mat_div, a.country, a.tp_year, tp_spread;
  quit;

  data dmimport.tactical_plan;
    set tactical_plan2;
  run;

%mend import_tactical_plan;