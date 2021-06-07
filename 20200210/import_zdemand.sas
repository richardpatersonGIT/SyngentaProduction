/***********************************************************************/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro IMPORT_ZDEMAND(zdemand_folder=, zdemand_type=);

  proc datasets lib=work nolist;
    delete zdemand_raw_all zdemand_raw;
  run;

  %read_folder(folder=%str(&zdemand_folder.), filelist=zdemand_list, ext_mask=txt, subfolders=Y);

  proc sql noprint;
    select count(*) into :zdemand_files_cnt from zdemand_list;
  quit;

  %do zdemand_file=1 %to &zdemand_files_cnt.;

    proc sql noprint;
      select path_file_name into :zdemand_file_name trimmed from zdemand_list where order=&zdemand_file.;
    quit;

    data ZDEMAND_RAW (drop=len name_start);
      infile "&zdemand_file_name." delimiter = '09'x missover dsd firstobs=2 lrecl=32767;
      length 
        Info_Str $4.
        Month $6.
        Week $6.
        Period $6.
        Div $2.
        Sales_Org $4.
        Sales_Off $4.
        Plant $4.
        _Variety $8.
        _Material $18.
        Material_Description $40.
        Sal $3.
        SRep $10.
        Sold_to_Party $10.
        _Confirmed_Sales_Forecast $16.
        _Returns_Planned $16.
        _Market_Uncertainty $16.
        _Confirmed_Sales_Plan $16.
        Base_Unit $3.
        filename $400.
        length filename_datetime $400.
        ;
      input 
        Info_Str $
        Month $
        Week $
        Period $
        Div $
        Sales_Org $
        Sales_Off $
        Plant $
        _Variety $
        _Material $
        Material_Description $
        Sal $
        SRep $
        Sold_to_Party $
        _Confirmed_Sales_Forecast $
        _Returns_Planned $
        _Market_Uncertainty $
        _Confirmed_Sales_Plan $
        Base_Unit $
        ;
        filename="&zdemand_file_name.";
        call scan(filename, -2, name_start , len, '_');
        filename_datetime=scan(substr(filename, name_start), 1, '.');
    run;

    data _null_;
      sleep=sleep(5);
    run;

    proc append base=ZDEMAND_RAW_ALL data=ZDEMAND_RAW;
    run;

  %end;

  data dmimport.ZDEMAND_&zdemand_type. (keep=
                                Info_Str
                                Month
                                Week
                                Period
                                Div
                                Sales_Org
                                Sales_Off
                                Plant
                                Variety
                                Material
                                Material_Description
/*                                Sal*/
/*                                SRep*/
/*                                Sold_to_Party*/
                                Confirmed_Sales_Forecast
/*                                Returns_Planned*/
/*                                Market_Uncertainty*/
                                Confirmed_Sales_Plan
                                Base_Unit
                                month_year 
                                month_month 
                                week_year 
                                week_week
                                filename
                                filename_datetime
                              );
    length month_year month_month week_year week_week material variety Confirmed_Sales_Forecast Returns_Planned Market_Uncertainty Confirmed_Sales_Plan 8.;
    set ZDEMAND_RAW_ALL;
    if input(month, 8.)^=0 then do;
      month_year=input(substr(month, 1, 4), 4.);
      month_month=input(substr(month, 5, 2), 2.);
    end;
    if input(week, 8.)^=0 then do;
      week_year=input(substr(week, 1, 4), 4.);
      week_week=input(substr(week, 5, 2), 2.);
    end;
    material=input(_material, 18.);
    variety=input(_variety, 8.);
    Confirmed_Sales_Forecast=input(_Confirmed_Sales_Forecast, best.);
    Returns_Planned=input(_Returns_Planned, best.);
    Market_Uncertainty=input(_Market_Uncertainty, best.);
    Confirmed_Sales_Plan=input(_Confirmed_Sales_Plan, best.);

    if ^(info_str='S967' and material<70000000) then output;
  run;

%mend IMPORT_ZDEMAND;

%macro PROCESS_ZDEMAND(zdemand_type=, zdemand_folder=);

  data ZDEMAND_&zdemand_type.(drop=rc)
       noseason_&zdemand_type.(drop=rc);
    length rc 8.;
    length season_week_start season_week_end season_start demand_date order_season 8.;
    length product_line_group $20. Product_Line $19. plg_rule $10. species $29. series $26.;
    length Region $3. Mat_div $2.;
    length sub_unit 8. DM_Process_stage $5. PF_for_sales_text $25. proc_stage $3. process_stage $5.;
    format season_start demand_date yymmdd10.;
    set dmimport.ZDEMAND_&zdemand_type.;

    if _n_=1 then do;
      declare hash zdemand_lookup(dataset: 'dmimport.zdemand_lookup');
        rc=zdemand_lookup.DefineKey ('Info_Str', 'Div', 'Sales_Org', 'Sales_off');
        rc=zdemand_lookup.DefineData ('Region', 'Mat_div');
        rc=zdemand_lookup.DefineDone();
      declare hash pmd_assortment(dataset: 'dmproc.PMD_assortment');
        rc=pmd_assortment.DefineKey ('region', 'variety');
        rc=pmd_assortment.DefineData ('season_week_start', 'season_week_end', 'product_line_group', 'product_line', 'plg_rule', 'species', 'series');
        rc=pmd_assortment.DefineDone();
      declare hash pmd_assortment_noregion(dataset: 'dmproc.PMD_assortment');
        rc=pmd_assortment_noregion.DefineKey ('variety');
        rc=pmd_assortment_noregion.DefineData ('season_week_start', 'season_week_end', 'product_line_group', 'product_line', 'plg_rule', 'species', 'series');
        rc=pmd_assortment_noregion.DefineDone();
      declare hash fps_assortment(dataset: 'dmimport.FPS_assortment');
        rc=fps_assortment.DefineKey ('material');
        rc=fps_assortment.DefineData ('sub_unit', 'DM_Process_stage', 'PF_for_sales_text');
        rc=fps_assortment.DefineDone();
      declare hash bi_assortment(dataset: 'dmimport.Material_class_table');
        rc=bi_assortment.DefineKey ('material', 'region');
        rc=bi_assortment.DefineData ('sub_unit', 'proc_stage');
        rc=bi_assortment.DefineDone();
    end;

    rc=zdemand_lookup.find();
    rc=pmd_assortment.find();
    if rc^=0 then do;
      rc=pmd_assortment_noregion.find();/*if variety is in wrong region take first found*/
    end;

    if ^missing(season_week_start) then do;
      if ^missing(month_month) then do;
        season_start=input(put(month_year, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
        demand_date=mdy(month_month, 15, month_year);
        demand_year=month_year;
      end;

      if ^missing(week_week) then do;
        season_start=input(put(week_year, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
        demand_date=input(put(week_year, 4.)||'W'||put(week_week, z2.)||'01', weekv9.);
        demand_year=week_year;
      end;

      if demand_date>=season_start then do;
        order_season=demand_year;
      end; else do;
        order_season=demand_year-1;
      end;
    end;

    If region = 'FPS' then do;
      rc=fps_assortment.find();
      process_stage=DM_Process_stage;
    end; else do;
      rc=bi_assortment.find();
      process_stage=proc_stage;
    end;

    if ^missing(order_season) then do;
      output ZDEMAND_&zdemand_type.;
    end; else do;
      output noseason_&zdemand_type.;
    end;

  run;

  proc sort data=ZDEMAND_&zdemand_type. out=zdemand_&zdemand_type._sorted;
    by region variety material sales_org sales_off plant demand_date order_season descending filename_datetime;
  run;

data dmproc.zdemand_&zdemand_type.
     dup_&zdemand_type.;
  set zdemand_&zdemand_type._sorted;
  by region variety material sales_org sales_off plant demand_date order_season descending filename_datetime;
  if first.order_season then output dmproc.zdemand_&zdemand_type.;
  else output dup_&zdemand_type.;
run;

%mend PROCESS_ZDEMAND;