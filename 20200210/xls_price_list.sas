/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %import_price_list(...)*/
/*Purpose: Imports Price List from folder (pl_folder=) with .xlsx extension*/
/*IN: price_list(1 file), extension=xlsx, sheet='Price List'*/
/*OUT: dmimport.price_list*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro import_price_list(pl_folder=);

  %read_folder(folder=%str(&pl_folder.), filelist=pl_files, ext_mask=xlsx);

  proc sql noprint;
  select path_file_name into :pl_file trimmed from pl_files where order=1;
  quit;

  PROC IMPORT OUT=price_list_raw
              DATAFILE="&pl_file."
              DBMS=  EXCELCS   REPLACE;
              sheet='Price list';
  RUN;

  data price_list(keep=region product_line species_code species_name material_division hash_mat_div price);
    set price_list_raw(rename=(price=_price));
    length hash_mat_div $3.;
    price=input(_price, best.);
    hash_mat_div=upcase(compress(compress(material_division, ' ', 'ka')));
  run;

  data dmimport.price_list;
    set price_list;
  run;

%mend import_price_list;