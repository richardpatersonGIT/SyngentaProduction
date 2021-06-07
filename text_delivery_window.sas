/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %IMPORT_DELIVERY_WINDOW(...)*/
/*Purpose: Imports Delivery window from folder (delivery_window_folder=) with .txt extension*/
/*IN: Delivery_window(multiple files), extension=txt, delimiter=;*/
/*OUT: dmimport.delivery_window*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro IMPORT_DELIVERY_WINDOW(delivery_window_folder=);

  %read_folder(folder=%str(&delivery_window_folder.), filelist=delivery_window_files, ext_mask=txt);

  proc datasets lib=work nolist;
    delete delivery_window_raw_all;
  run;

  proc sql noprint;
    select count(*) into :delivery_window_files_cnt from delivery_window_files;
  quit;

  %do dw_file=1 %to &delivery_window_files_cnt.;

    proc sql noprint;
      select path_file_name into :delivery_window_file trimmed from delivery_window_files where order = &dw_file.;
    quit;

    data delivery_window_raw;
      infile "&delivery_window_file." delimiter = ';' MISSOVER DSD lrecl=32767  firstobs=2;
      length Species $4.
            Species_Desc $29.
            _Variety $8.
            Variety_Desc $30.
            _Material $8.
            Material_Description $40.
            Plant $4.
            Process_stage $3.
            Process_stage_Desc $20.
            PLC_code $2.
            Material_Group $6.
            Delivery_Week $6.
            Order_Week $6.
            URC_Week $6.
            Qty_in_Box $7.
            Qty_in_Pc $11.
            Open_Closed_week_status $4.
            Requested_EAME_Sales $3.
            Confirmed_EAME_Sales $3.
            Unconfirmed_EAME_Sales $1.
            Constrained_Demand $1.;

    input Species
          Species_Desc
          _Variety
          Variety_Desc
          _Material
          Material_Description
          Plant
          Process_stage
          Process_stage_Desc
          PLC_code
          Material_Group
          Delivery_Week
          Order_Week
          URC_Week
          Qty_in_Box
          Qty_in_Pc
          Open_Closed_week_status
          Requested_EAME_Sales
          Confirmed_EAME_Sales
          Unconfirmed_EAME_Sales
          Constrained_Demand
          ;
    run;

    data _null_;
      sleep=sleep(0);
    run;

    proc append base=delivery_window_raw_all data=delivery_window_raw;
    run;

  %end;

  data delivery_window_original(keep=variety material delivery_week_year delivery_week_week status);
    set delivery_window_raw_all;
    length delivery_week_year delivery_week_week 8.;
    material=input(_material, 8.);
    variety=input(_variety, 8.);
    delivery_week_year=input(substr(put(strip(delivery_week), 6.), 1, 4), 8.);
    delivery_week_week=input(substr(put(strip(delivery_week), 6.), 5, 2), 8.);
    if open_closed_week_status = 'CLSD' then status=0;
    if open_closed_week_status = 'APPR' then status=1;
  run;

  proc sql;
    create table delivery_window_to_open as
      select distinct a.material, 1 as status
       from (select distinct material from delivery_window_original) a
       left join (select distinct material from delivery_window_original where status=0) b on a.material=b.material
       left join (select distinct material from delivery_window_original where status=1) c on a.material=c.material
       where ^missing(b.material) and missing(c.material);
  quit;

  proc sql;
    create table delivery_window_updated as
    select 
      a.variety,
      a.material,
      a.delivery_week_year,
      a.delivery_week_week,
      coalesce(b.status, a.status) as status
    from delivery_window_original a
    left join delivery_window_to_open b on a.material=b.material;
  quit;

  data dmimport.delivery_window;
    set delivery_window_updated;
  run;


%mend IMPORT_DELIVERY_WINDOW;