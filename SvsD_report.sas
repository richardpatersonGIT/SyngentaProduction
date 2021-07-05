/***********************************************************************/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\cleanup_xlsx_bak_folder.sas";
%include "&sas_applications_folder.\filter_orders.sas";


%macro SvsD_read_history(history_suffix=,
                         season_year=, 
                         season_week=,
                         filter_season=,
                         filter_week=);

  data _null_;
    sleep=sleep(0);
  run;

  proc datasets library=work nolist;
    delete SvsD_history;
  run;

  proc contents data=shw._all_ out=shw_contents noprint;
  run;

  proc sql noprint;
    create table shw_contents1 as
    select distinct memname as dsname from shw_contents where upper(memname) like 'SHW_%' order by dsname;
  quit;

  data filelist2;
    set shw_contents1;
    year=input(substr(dsname, 5, 4), 4.);
    week=input(compress(substr(dsname, 12, 2),, "kd"),  8.);
    if (year=&season_year. and week=&season_week.) then output;
  run;

  data filelist3;
    set filelist2;
    order=_n_;
  run;

  proc sql noprint;
  select count(*) into : flcnt trimmed from filelist3;
  quit;
  
  %if "&flcnt."="0" %then %do;
    data SvsD_history1;
    length region $3. variety 8. material 8. order_season 8. cnf_qty 8. historical_sales 8.;
    run;
  %end; %else %do;
    %do i=1 %to &flcnt.;
      proc sql noprint;
        select dsname, year, week into :dsname trimmed, :year trimmed, :week trimmed from filelist3 where order = &i.;
      quit;

      data SvsD_history_raw(keep=soldto_nr sls_org sls_off shipto_cntry material variety SchedLine_Cnf_deldte cnf_qty order_type mat_div rsn_rej_cd filename);
        length filename $200.;
        set shw.&dsname.;
        filename="&dsname.";
      run;

      data SvsD_history_tmp(drop=sls_org sls_off shipto_cntry order_type rsn_rej_cd);
        set SvsD_history_raw;
        length unique_code $10. year 8. week 8.;
        unique_code=cats(sls_org, sls_off, shipto_cntry);
        year=&year.;
        week=&week.;
        if ((mat_div in ('6B', '6C') and order_type in ('ZYPD', 'ZFD1', 'ZYPL', 'ZMTO')) or (mat_div='6A' and order_type in ('YQOR', 'ZMTO'))) and missing(rsn_rej_cd) and ^missing(SchedLine_Cnf_deldte) then output;
      run;

      /*<corection of date format>*/
      proc contents data=SvsD_history_tmp out=contents noprint;
      run;

      %let datetype=1;

      proc sql noprint;
        select type into :datetype trimmed from contents where lower(name)="schedline_cnf_deldte";
      quit;

      %if "&datetype."="2" %then %do;
        data SvsD_history_tmp(drop=_SchedLine_Cnf_deldte);
          set SvsD_history_tmp(rename=(SchedLine_Cnf_deldte=_SchedLine_Cnf_deldte));
          SchedLine_Cnf_deldte=input(_SchedLine_Cnf_deldte, 8.);
          SchedLine_Cnf_deldte=SchedLine_Cnf_deldte-21916;  
        run;
      %end;
      /*</corection of date format>*/

      proc datasets lib=work memtype=data nolist;
         modify SvsD_history_tmp;
         attrib _all_ format=;
      run;
      
    data _null_;
        sleep=sleep(0);
      run;

      proc append base=SvsD_history data=SvsD_history_tmp;
      run;

    %end;

    data SvsD_history1(drop=rc);
      set SvsD_history;
      length region $3. territory $3. country $6.;
      length sub_unit 8.;
      length  season_week_start season_week_end 
              Order_season_start order_year order_season order_week Order_yweek order_month 8.;
      length hash_species_name $29. product_line_group $20.;
      format Order_season_start yymmdd10.;

      if _n_=1 then do;
        declare hash cl(dataset: 'dmimport.Country_lookup');
          rc=cl.DefineKey ('unique_code');
          rc=cl.DefineData ('region', 'territory', 'country');
          rc=cl.DefineDone();
        declare hash stl(dataset: 'dmimport.Soldto_nr_lookup');
          rc=stl.DefineKey ('soldto_nr');
          rc=stl.DefineData ('region', 'territory', 'country');
          rc=stl.DefineDone();
        declare hash material_assortment(dataset: 'dmproc.material_assortment');
          rc=material_assortment.DefineKey ('region', 'material');
          rc=material_assortment.DefineData ('sub_unit');
          rc=material_assortment.DefineDone();          
        declare hash pmd_assortment(dataset: 'dmproc.PMD_assortment');
          rc=pmd_assortment.DefineKey ('region', 'variety');
          rc=pmd_assortment.DefineData ('season_week_start', 'season_week_end', 'hash_species_name', 'product_line_group');
          rc=pmd_assortment.DefineDone();
        declare hash pmd_assortment_noregion(dataset: 'dmproc.PMD_assortment');
          rc=pmd_assortment_noregion.DefineKey ('variety');
          rc=pmd_assortment_noregion.DefineData ('season_week_start', 'season_week_end', 'hash_species_name', 'product_line_group');
          rc=pmd_assortment_noregion.DefineDone();
      end;

      rc=cl.find(); /*gets territory and country from country_lookup*/
      if region='BI' then do;
        rc=stl.find();/*gets territory and country from soldto_nr_lookup (if found overwrite the country_lookup)*/
      end;

      rc=material_assortment.find();

      if ^missing(sub_unit) then do;
        historical_sales=cnf_qty * sub_unit;
      end;

      rc=pmd_assortment.find(); 
      if rc^=0 then do;
        rc=pmd_assortment_noregion.find(); 
      end;

      order_year=year(SchedLine_Cnf_deldte);
      order_month=month(SchedLine_Cnf_deldte);
      if ^missing(season_week_start) then do;
        order_season_start=input(put(order_year, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
        if SchedLine_Cnf_deldte >= order_season_start then do;
          order_season=order_year; 
        end; else do;
          order_season=order_year-1;
        end;
      end;

      order_yweek=input(substr(put(SchedLine_Cnf_deldte, weekv9.), 1, 4), 4.);
      order_week=input(substr(put(SchedLine_Cnf_deldte, weekv9.), 6, 2), 2.);
      if order_week=53 then order_week=52;

      if ^missing(region) and ^missing(material) and ^missing(variety) and order_season=&filter_season.
      %if ^("&filter_week."="" or "&filter_week."=".") %then %do;
        and season_week_end = &filter_week. 
      %end;
      then output;
    run;

  %end;

  %filter_orders(in_table=SvsD_history1, out_table=SvsD_history2);

  proc sql;
    create table SvsD_history_&history_suffix. as
    select region, variety, material,
           sum(historical_sales) as history_plant_&history_suffix., sum(cnf_qty) as history_box_&history_suffix.
    from SvsD_history2
    group by region, variety, material;
  quit;

%mend SvsD_read_history;

%macro read_svsd_metadata();

  PROC IMPORT OUT=svsd_report_md_raw
              DATAFILE="&sd_metadata_file."
              DBMS=  EXCELCS  REPLACE; 
  RUN;

  data svsd_report_md_raw1;
    set svsd_report_md_raw;
    if ^missing(coalesceC(of _character_)) or ^missing(coalesce(of _numeric_)) then output;
  run;

  data dmimport.svsd_report_md(drop=_: char_report_week rename=(season=order_season product_form=PF_for_sales_text));
    length hash 8. season_week_start 8. report_year 8. report_week 8. char_report_week $6.;
    set svsd_report_md_raw1(rename=(season_week_start=_season_week_start
                                    report_week=_report_week
                                    ));
    hash=1;
    season_week_start=input(_season_week_start, best.);
    char_report_week=input(_report_week, best.);
    report_year=put(substr(char_report_week, 1, 4), 4.);
    report_week=put(substr(char_report_week, 5, 2), 2.);
  run;

%mend read_svsd_metadata;

%macro REPORT_SALES_VS_DEMAND();

  %read_svsd_metadata();

  proc sql noprint;
    select count(*) into :region trimmed from dmimport.svsd_report_md where ^missing(region);
    select count(*) into :product_line trimmed from dmimport.svsd_report_md where ^missing(product_line);
    select count(*) into :species trimmed from dmimport.svsd_report_md where ^missing(species);
    select count(*) into :series trimmed from dmimport.svsd_report_md where ^missing(series);
    select count(*) into :variety trimmed from dmimport.svsd_report_md where ^missing(variety);
    select count(*) into :material trimmed from dmimport.svsd_report_md where ^missing(material);
    select count(*) into :PF_for_sales_text trimmed from dmimport.svsd_report_md where ^missing(PF_for_sales_text);
    select count(*) into :process_stage trimmed from dmimport.svsd_report_md where ^missing(process_stage);
    select count(*) into :mat_div trimmed from dmimport.svsd_report_md where ^missing(mat_div);
    select count(*) into :product_line_group trimmed from dmimport.svsd_report_md where ^missing(product_line_group);
    select count(*) into :season_week_start trimmed from dmimport.svsd_report_md where ^missing(season_week_start);
  quit;

  proc sql noprint;
    select max(order_season) into :current_season trimmed from dmimport.svsd_report_md where ^missing(order_season);
    select max(report_year) into :report_year trimmed from dmimport.svsd_report_md where ^missing(report_year);
    select max(report_week) into :report_week trimmed from dmimport.svsd_report_md where ^missing(report_week);
    select max(unit_for_ypl) into :unit_for_ypl trimmed from dmimport.svsd_report_md where ^missing(unit_for_ypl);
  quit;

  %let s_3=%eval(&current_season.-3);
  %let s_2=%eval(&current_season.-2);
  %let s_1=%eval(&current_season.-1);
  %let s0=&current_season.;
  %let s1=%eval(&current_season.+1);
  %let s2=%eval(&current_season.+2);
  %let s3=%eval(&current_season.+3);

  proc sql noprint;
    create table SvsD_season_ends as
    select season_week_end from dmimport.seasons_general where ^missing(season_week_end)
    union select season_week_end from dmimport.seasons_grouping_code_exc where ^missing(season_week_end)
    union select season_week_end from dmimport.seasons_species_exc where ^missing(season_week_end);
  quit;
  data SvsD_history_list(drop=season_week_end);
    length history_suffix $10.;
    length season_year 8. season_week 8.;
    length filter_season 8. filter_week 8.;
    set SvsD_season_ends;
/*
    if _n_=1 then do;

      history_suffix="cw_s_1"; 
      filter_season=&s_1.;
      if 
      season_year=&s_1.;
      season_week=&report_week.;
      call missing(filter_week);
      output;

    end;
*/
      history_suffix="cw_s_1";
      filter_season=&s_1.;
      filter_week=season_week_end;
      if season_week_end=52 then do;
        season_year=&s_1.;
      end; else do;
        season_year=&s_1.+1;
      end;
      season_week=&report_week.;
      output;

      history_suffix="eos_s_1";
      filter_season=&s_1.;
      filter_week=season_week_end;
      if season_week_end=52 then do;
        season_year=&s_1.;
      end; else do;
        season_year=&s_1.+1;
      end;
      season_week=season_week_end;
      output;

      history_suffix="eos_s_2";
      filter_season=&s_2.;
      if season_week_end=52 then do;
        season_year=&s_2.;
      end; else do;
        season_year=&s_2.+1;
      end;
      season_week=season_week_end;
      output;

    history_suffix="eos_s_3";
      filter_season=&s_3.;
      if season_week_end=52 then do;
        season_year=&s_3.;
      end; else do;
        season_year=&s_3.+1;
      end;
      season_week=season_week_end;
      output;
  run;

  data SvsD_history_list;
    set SvsD_history_list;
    hid=_n_;
  run;

  data svsd_history_all;
    length region $3. variety 8. material 8.
           history_plant_cw_s_1 8.  history_box_cw_s_1 8. 
           history_plant_eos_s_1 8. history_box_eos_s_1 8.
           history_plant_eos_s_2 8. history_box_eos_s_2 8. 
           history_plant_eos_s_3 8. history_box_eos_s_3 8. ;
  run;

  proc sql noprint;
    select count(*) into: history_cnt from SvsD_history_list;
  quit;

  %do hid=1 %to &history_cnt.;

    proc sql noprint;
      select history_suffix, season_year, season_week, filter_season, filter_week
        into :history_suffix trimmed, :season_year trimmed, :season_week trimmed, :filter_season trimmed, :filter_week trimmed
      from SvsD_history_list where hid=&hid.;
    quit;

    %SvsD_read_history(history_suffix=&history_suffix., season_year=&season_year., season_week=&season_week., filter_season=&filter_season., filter_week=&filter_week.);

    data svsd_history_all;
      merge svsd_history_all SvsD_history_&history_suffix.;
      by region variety material;
      if ^missing(region) then output;
    run;

  %end;

  proc sql;
    create table zdemand_sap_zdemand_aggr_s0 as
    select region, variety, material, 
           sum(Confirmed_Sales_Forecast) as sap_normal_box_s0, 
           sum(Confirmed_Sales_Plan) as sap_constrained_box_s0
    from dmproc.zdemand_sap_zdemand
    where order_season=&s0.
    group by region, variety, material;
  quit;

  proc sql;
    create table zdemand_sap_zdemand_aggr_s1 as
    select region, variety, material, 
           sum(Confirmed_Sales_Forecast) as sap_normal_box_s1, 
           sum(Confirmed_Sales_Plan) as sap_constrained_box_s1
    from dmproc.zdemand_sap_zdemand
    where order_season=&s1.
    group by region, variety, material;
  quit;

  proc sql;
    create table zdemand_sap_zdemand_aggr_s2 as
    select region, variety, material, 
           sum(Confirmed_Sales_Forecast) as sap_normal_box_s2, 
           sum(Confirmed_Sales_Plan) as sap_constrained_box_s2
    from dmproc.zdemand_sap_zdemand
    where order_season=&s2.
    group by region, variety, material;
  quit;

  proc sql;
    create table zdemand_sap_zdemand_aggr_s3 as
    select region, variety, material, 
           sum(Confirmed_Sales_Forecast) as sap_normal_box_s3, 
           sum(Confirmed_Sales_Plan) as sap_constrained_box_s3
    from dmproc.zdemand_sap_zdemand
    where order_season=&s3.
    group by region, variety, material;
  quit;

  proc sql;
    create table zdemand_so_demand_aggr_s0 as
    select region, variety, material, 
           sum(Confirmed_Sales_Forecast) as so_normal_box_s0, 
           sum(Confirmed_Sales_Plan) as so_constrained_box_s0
    from  dmproc.zdemand_so_demand
    where order_season=&s0.
    group by region, variety, material;
  quit;

  proc sql;
    create table zdemand_so_demand_aggr_s1 as
    select region, variety, material, 
           sum(Confirmed_Sales_Forecast) as so_normal_box_s1, 
           sum(Confirmed_Sales_Plan) as so_constrained_box_s1
    from  dmproc.zdemand_so_demand
    where order_season=&s1.
    group by region, variety, material;
  quit;

  proc sql;
    create table zdemand_so_demand_aggr_s2 as
    select region, variety, material, 
           sum(Confirmed_Sales_Forecast) as so_normal_box_s2, 
           sum(Confirmed_Sales_Plan) as so_constrained_box_s2
    from  dmproc.zdemand_so_demand
    where order_season=&s2.
    group by region, variety, material;
  quit;

  proc sql;
    create table zdemand_so_demand_aggr_s3 as
    select region, variety, material, 
           sum(Confirmed_Sales_Forecast) as so_normal_box_s3, 
           sum(Confirmed_Sales_Plan) as so_constrained_box_s3
    from  dmproc.zdemand_so_demand
    where order_season=&s3.
    group by region, variety, material;
  quit;

  %filter_orders(in_table=dmproc.orders_all, out_table=orders_filtered);

  proc sql;
    create table inv_orders_aggr as
    select region, variety, material,
           sum(historical_sales) as inv_orders_demand_plant, 
           sum(cnf_qty) as inv_orders_demand_box
    from orders_filtered 
    where order_season=&s0. and ((order_year=&report_year. and order_week<=&report_week.) or (order_year<&report_year.)) 
    group by region, variety, material;
  quit;

  proc sql;
    create table ooh_orders_aggr as
    select region, variety, material,
           sum(historical_sales) as ooh_orders_demand_plant, 
           sum(cnf_qty) as ooh_orders_demand_box
    from orders_filtered
    where order_season=&s0. and ((order_year=&report_year. and order_week>&report_week.) or (order_year>&report_year.))
    group by region, variety, material;
  quit;


  data demand_all;
    merge inv_orders_aggr
          ooh_orders_aggr
          zdemand_sap_zdemand_aggr_s0
          zdemand_sap_zdemand_aggr_s1
          zdemand_sap_zdemand_aggr_s2
          zdemand_sap_zdemand_aggr_s3
          zdemand_so_demand_aggr_s0 
          zdemand_so_demand_aggr_s1
          zdemand_so_demand_aggr_s2
          zdemand_so_demand_aggr_s3
          ;
     by region variety material;
     order_season=&s0.;
  run;

  /*keeping all history for extrapolation, filtering out afterwards*/
  data SvsD_demand;
    merge demand_all svsd_history_all;
    by region variety material;
  run;

  data SvsD_variety(drop=rc);
      set SvsD_demand;
      length  season_week_start 8.;
      length product_line_group $20. Product_Line $19. species $29. series $26. variety_name $60. 
                     variety_plc_current $2. variety_plc_future $2. variety_plc_future_date 8. 
                     variety_global_plc_current $2.;
      format variety_plc_future_date yymmdd10.;
      if _n_=1 then do;
        declare hash pmd_assortment(dataset: 'dmproc.PMD_assortment(rename=(Current_PLC=variety_plc_current 
                                                                            Future_PLC=variety_plc_future 
                                                                            Future_PLC_Active_date=variety_plc_future_date
                                                                            Global_Current_PLC=variety_global_plc_current))');
          rc=pmd_assortment.DefineKey ('region', 'variety');
          rc=pmd_assortment.DefineData ('season_week_start', 'product_line_group', 'product_line', 
                                        'species', 'series', 'variety_name', 'variety_plc_current', 
                                        'variety_plc_future', 'variety_plc_future_date', 
                                        'variety_global_plc_current');
          rc=pmd_assortment.DefineDone();
      end;

      rc=pmd_assortment.find(); 
      *if region^='BI' then call missing(variety_global_plc_current);
  run;

  data SvsD_material(drop=rc fps_material_name bi_material_name);
    set SvsD_variety;
    length sub_unit 8.  product_form $25. process_stage $5. material_division $2. 
           material_plc_current $2. material_plc_future $2. material_plc_future_date 8. fps_material_name $40. bi_material_name $40.
           material_name $40.;
    format material_plc_future_date yymmdd10.;

    if _n_=1 then do;
      declare hash material_assortment(dataset: 'dmproc.material_assortment');
        rc=material_assortment.DefineKey ('region', 'material');
        rc=material_assortment.DefineData ('sub_unit', 'process_stage', 'product_form', 'material_division', 
                                      'material_plc_current', 'material_plc_future', 'material_plc_future_date', 'fps_material_name', 'bi_material_name');
        rc=material_assortment.DefineDone(); 
    end;

    rc=material_assortment.find();
    material_name=coalescec(fps_material_name, bi_material_name); 

  run;

  data SvsD_filtered(
    drop=  rc 
          hash 
          reject 
          );
    length rc hash reject 8.;
    set SvsD_material;
    reject=0;

    %if &region. > 0 %then %do;
      if _n_=1 then do;
        declare hash h_region(dataset: 'dmimport.svsd_report_md(where=(^missing(region)))');
          rc=h_region.DefineKey ('region');
          rc=h_region.DefineData ('hash');
          rc=h_region.DefineDone();
      end;
      hash=0;
      rc=h_region.find();
      if hash=0 then reject=1;
    %end;

    %if &product_line. > 0 %then %do;
      if _n_=1 then do;
        declare hash h_product_line(dataset: 'dmimport.svsd_report_md(where=(^missing(product_line)))');
          rc=h_product_line.DefineKey ('product_line');
          rc=h_product_line.DefineData ('hash');
          rc=h_product_line.DefineDone();
      end;
      hash=0;
      rc=h_product_line.find();
      if hash=0 then reject=1;
    %end;

    %if &species. > 0 %then %do;
      if _n_=1 then do;
        declare hash h_species(dataset: 'dmimport.svsd_report_md(where=(^missing(species)))');
          rc=h_species.DefineKey ('species');
          rc=h_species.DefineData ('hash');
          rc=h_species.DefineDone();
      end;
      hash=0;
      rc=h_species.find();
      if hash=0 then reject=1;
    %end;

    %if &series. > 0 %then %do;
      if _n_=1 then do;
        declare hash h_series(dataset: 'dmimport.svsd_report_md(where=(^missing(series)))');
          rc=h_series.DefineKey ('series');
          rc=h_series.DefineData ('hash');
          rc=h_series.DefineDone();
      end;
      hash=0;
      rc=h_series.find();
      if hash=0 then reject=1;
    %end;

    %if &variety. > 0 %then %do;
      if _n_=1 then do;
        declare hash h_variety(dataset: 'dmimport.svsd_report_md(where=(^missing(variety)))');
          rc=h_variety.DefineKey ('variety');
          rc=h_variety.DefineData ('hash');
          rc=h_variety.DefineDone();
      end;
      hash=0;
      rc=h_variety.find();
      if hash=0 then reject=1;
    %end;

    %if &material. > 0 %then %do;
      if _n_=1 then do;
        declare hash h_material(dataset: 'dmimport.svsd_report_md(where=(^missing(material)))');
          rc=h_material.DefineKey ('material');
          rc=h_material.DefineData ('hash');
          rc=h_material.DefineDone();
      end;
      hash=0;
      rc=h_material.find();
      if hash=0 then reject=1;
    %end;

    %if &PF_for_sales_text. > 0 %then %do;
      if _n_=1 then do;
        declare hash h_product_form(dataset: 'dmimport.svsd_report_md(where=(^missing(PF_for_sales_text)))');
          rc=h_product_form.DefineKey ('PF_for_sales_text');
          rc=h_product_form.DefineData ('hash');
          rc=h_product_form.DefineDone();
      end;
      hash=0;
      rc=h_product_form.find();
      if hash=0 then reject=1;
    %end;

    %if &process_stage. > 0 %then %do;
      if _n_=1 then do;
        declare hash h_process_stage(dataset: 'dmimport.svsd_report_md(where=(^missing(process_stage)))');
          rc=h_process_stage.DefineKey ('process_stage');
          rc=h_process_stage.DefineData ('hash');
          rc=h_process_stage.DefineDone();
      end;
      hash=0;
      rc=h_process_stage.find();
      if hash=0 then reject=1;
    %end;

    %if &mat_div. > 0 %then %do;
      if _n_=1 then do;
        declare hash h_mat_div(dataset: 'dmimport.svsd_report_md(where=(^missing(mat_div)))');
          rc=h_mat_div.DefineKey ('mat_div');
          rc=h_mat_div.DefineData ('hash');
          rc=h_mat_div.DefineDone();
      end;
      hash=0;
      rc=h_mat_div.find();
      if hash=0 then reject=1;
    %end;

    %if &product_line_group. > 0 %then %do;
      if _n_=1 then do;
        declare hash h_product_line_group(dataset: 'dmimport.svsd_report_md(where=(^missing(product_line_group)))');
          rc=h_product_line_group.DefineKey ('product_line_group');
          rc=h_product_line_group.DefineData ('hash');
          rc=h_product_line_group.DefineDone();
      end;
      hash=0;
      rc=h_product_line_group.find();
      if hash=0 then reject=1;
    %end;


    %if &season_week_start. > 0 %then %do;
      if _n_=1 then do;
        declare hash h_season_week_start(dataset: 'dmimport.svsd_report_md(where=(^missing(season_week_start)))');
          rc=h_season_week_start.DefineKey ('season_week_start');
          rc=h_season_week_start.DefineData ('hash');
          rc=h_season_week_start.DefineDone();
      end;
      hash=0;
      rc=h_season_week_start.find();
      if hash=0 then reject=1;
    %end;

    if reject=0 then output;
  run;

  proc sql;
    create table extrapolation_s_1 as
    select region, product_line_group, species, material_division, 
           sum(history_plant_cw_s_1) as series_sum_cw, 
           sum(history_plant_eos_s_1) as series_sum_eos,
           case 
           when calculated series_sum_eos>0 then calculated series_sum_cw/calculated series_sum_eos
           else . 
           end as extrapolation
    from SvsD_filtered
    group by region, product_line_group, species, material_division;
  quit;

  proc sql;
    create table SvsD_extrapolation as
    select a.*, 
           b.extrapolation
    from SvsD_filtered a
    left join extrapolation_s_1 b on a.region=b.region and a.product_line_group=b.product_line_group and a.species=b.species and a.material_division=b.material_division;
  quit;

  /*filter out only materials that are in orders, zdemand or signed off demand*/
  proc sql;
    create table SvsD_extrapolation_filter as
    select a.*
    from SvsD_extrapolation a
    right join demand_all b on a.region=b.region and a.variety=b.variety and a.material=b.material
    where ^missing(a.region);
  quit;

  data SvsD_calculation;
    set SvsD_extrapolation_filter;
    length item_measurement $3.;
    format extrapolation adj_extrapolation prc_sap_so prc_ext_sap prc_ext_so percentn8.2;
    format extrapolated_orders_plant extrapolated_orders_box 18.0;
    if missing(extrapolation) or extrapolation>1 then do;
      adj_extrapolation=1;
    end; else do;
      adj_extrapolation=extrapolation;
    end;
    total_orders_demand_plant=coalesce(inv_orders_demand_plant, 0)+coalesce(ooh_orders_demand_plant, 0);
    total_orders_demand_box=coalesce(inv_orders_demand_box, 0)+coalesce(ooh_orders_demand_box, 0);
    extrapolated_orders_plant=total_orders_demand_plant/adj_extrapolation;
    extrapolated_orders_box=total_orders_demand_box/adj_extrapolation;
    if material_division='6A' then item_measurement='KS';
    %if "&unit_for_ypl."="BOX" %then %do;
      if material_division='6B' then item_measurement='BOX';
    %end; %else %do;
      if material_division='6B' then item_measurement='YPL';
    %end;
    if material_division='6C' then item_measurement='URC';

    if material_division='6B' then do; /*6B*/
        diff_sap_so_box=coalesce(sap_normal_box_s0, 0)-coalesce(so_normal_box_s0, 0);
        if so_normal_box_s0>0 then prc_sap_so_box=sap_normal_box_s0/so_normal_box_s0-1;
        diff_ext_sap_box=coalesce(extrapolated_orders_box, 0)-coalesce(sap_normal_box_s0, 0);
        if sap_normal_box_s0>0 then prc_ext_sap_box=extrapolated_orders_box/sap_normal_box_s0-1;
        diff_ext_so_box=coalesce(extrapolated_orders_box, 0)-coalesce(so_normal_box_s0, 0);
        if so_normal_box_s0>0 then prc_ext_so_box=extrapolated_orders_box/so_normal_box_s0-1;
        unconsumed_demand_box=coalesce(sap_constrained_box_s0, 0)-total_orders_demand_box;

        diff_sap_so_plant=coalesce(sap_normal_box_s0, 0)*sub_unit-coalesce(so_normal_box_s0, 0)*sub_unit;
        if so_normal_box_s0>0 then prc_sap_so_plant=((sap_normal_box_s0*sub_unit)/(so_normal_box_s0*sub_unit))-1;
        diff_ext_sap_plant=coalesce(extrapolated_orders_plant, 0)-coalesce(sap_normal_box_s0, 0)*sub_unit;
        if sap_normal_box_s0>0 then prc_ext_sap_plant=extrapolated_orders_plant/(sap_normal_box_s0*sub_unit)-1;
        diff_ext_so_plant=coalesce(extrapolated_orders_plant, 0)-(coalesce(so_normal_box_s0, 0)*sub_unit);
        if so_normal_box_s0>0 then prc_ext_so_plant=(extrapolated_orders_plant/(so_normal_box_s0*sub_unit))-1;
        unconsumed_demand_plant=(coalesce(sap_constrained_box_s0, 0)*sub_unit)-total_orders_demand_plant;
    end; else do;/*6A, 6C*/
        diff_sap_so_plant=coalesce(sap_normal_box_s0, 0)-coalesce(so_normal_box_s0, 0);
        if so_normal_box_s0>0 then prc_sap_so_plant=sap_normal_box_s0/so_normal_box_s0-1;
        diff_ext_sap_plant=coalesce(extrapolated_orders_plant, 0)-coalesce(sap_normal_box_s0, 0);
        if sap_normal_box_s0>0 then prc_ext_sap_plant=extrapolated_orders_plant/sap_normal_box_s0-1;
        diff_ext_so_plant=coalesce(extrapolated_orders_plant, 0)-coalesce(so_normal_box_s0, 0);
        if so_normal_box_s0>0 then prc_ext_so_plant=extrapolated_orders_plant/so_normal_box_s0-1;
        unconsumed_demand_plant=coalesce(sap_constrained_box_s0, 0)-total_orders_demand_plant;
    end;
  run;

  data SvsD_units;
    set SvsD_calculation;
    if material_division='6B' then do; 
      %if "&unit_for_ypl."="BOX" %then %do;/*BOX - 6B*/
        inv_orders=inv_orders_demand_box;
        ooh_orders=ooh_orders_demand_box;
        sap_normal_s0=sap_normal_box_s0;
        sap_normal_s1=sap_normal_box_s1;
        sap_normal_s2=sap_normal_box_s2;
        sap_normal_s3=sap_normal_box_s3;
        so_normal_s0=so_normal_box_s0;
        so_normal_s1=so_normal_box_s1;
        so_normal_s2=so_normal_box_s2;
        so_normal_s3=so_normal_box_s3;
        sap_constrained_s0=sap_constrained_box_s0;
        sap_constrained_s1=sap_constrained_box_s1;
        sap_constrained_s2=sap_constrained_box_s2;
        sap_constrained_s3=sap_constrained_box_s3;
        so_constrained_s0=so_constrained_box_s0;
        so_constrained_s1=so_constrained_box_s1;
        so_constrained_s2=so_constrained_box_s2;
        so_constrained_s3=so_constrained_box_s3;
        s_1_cw_history=history_box_cw_s_1;
        s_1_eos_history=history_box_eos_s_1;
        s_2_eos_history=history_box_eos_s_2;
        s_3_eos_history=history_box_eos_s_3;
        total_orders=total_orders_demand_box;
        extrapolated_orders=extrapolated_orders_box;
        unconsumed_demand=unconsumed_demand_box;
        diff_sap_so=diff_sap_so_box;
        diff_ext_sap=diff_ext_sap_box;
        diff_ext_so=diff_ext_so_box;
        prc_sap_so=prc_sap_so_box;
        prc_ext_sap=prc_ext_sap_box;
        prc_ext_so=prc_ext_so_box;
      %end; %else %do;/*NOT_BOX - 6B*/
        inv_orders=inv_orders_demand_plant;
        ooh_orders=ooh_orders_demand_plant;
        sap_normal_s0=sap_normal_box_s0*sub_unit;
        sap_normal_s1=sap_normal_box_s1*sub_unit;
        sap_normal_s2=sap_normal_box_s2*sub_unit;
        sap_normal_s3=sap_normal_box_s3*sub_unit;
        so_normal_s0=so_normal_box_s0*sub_unit;
        so_normal_s1=so_normal_box_s1*sub_unit;
        so_normal_s2=so_normal_box_s2*sub_unit;
        so_normal_s3=so_normal_box_s3*sub_unit;
        sap_constrained_s0=sap_constrained_box_s0*sub_unit;
        sap_constrained_s1=sap_constrained_box_s1*sub_unit;
        sap_constrained_s2=sap_constrained_box_s2*sub_unit;
        sap_constrained_s3=sap_constrained_box_s3*sub_unit;
        so_constrained_s0=so_constrained_box_s0*sub_unit;
        so_constrained_s1=so_constrained_box_s1*sub_unit;
        so_constrained_s2=so_constrained_box_s2*sub_unit;
        so_constrained_s3=so_constrained_box_s3*sub_unit;
        s_1_cw_history=history_plant_cw_s_1;
        s_1_eos_history=history_plant_eos_s_1;
        s_2_eos_history=history_plant_eos_s_2;
        s_3_eos_history=history_plant_eos_s_3;
        total_orders=total_orders_demand_plant;
        extrapolated_orders=extrapolated_orders_plant;
        unconsumed_demand=unconsumed_demand_plant;
        diff_sap_so=diff_sap_so_plant;
        diff_ext_sap=diff_ext_sap_plant;
        diff_ext_so=diff_ext_so_plant;
        prc_sap_so=prc_sap_so_plant;
        prc_ext_sap=prc_ext_sap_plant;
        prc_ext_so=prc_ext_so_plant;
      %end;
    end; else do;/*6A, 6C*/
        inv_orders=inv_orders_demand_plant;
        ooh_orders=ooh_orders_demand_plant;
        sap_normal_s0=sap_normal_box_s0;
        sap_normal_s1=sap_normal_box_s1;
        sap_normal_s2=sap_normal_box_s2;
        sap_normal_s3=sap_normal_box_s3;
        so_normal_s0=so_normal_box_s0;
        so_normal_s1=so_normal_box_s1;
        so_normal_s2=so_normal_box_s2;
        so_normal_s3=so_normal_box_s3;
        sap_constrained_s0=sap_constrained_box_s0;
        sap_constrained_s1=sap_constrained_box_s1;
        sap_constrained_s2=sap_constrained_box_s2;
        sap_constrained_s3=sap_constrained_box_s3;
        so_constrained_s0=so_constrained_box_s0;
        so_constrained_s1=so_constrained_box_s1;
        so_constrained_s2=so_constrained_box_s2;
        so_constrained_s3=so_constrained_box_s3;
        s_1_cw_history=history_plant_cw_s_1;
        s_1_eos_history=history_plant_eos_s_1;
        s_2_eos_history=history_plant_eos_s_2;
        s_3_eos_history=history_plant_eos_s_3;
        total_orders=total_orders_demand_plant;
        extrapolated_orders=extrapolated_orders_plant;
        unconsumed_demand=unconsumed_demand_plant;
        diff_sap_so=diff_sap_so_plant;
        diff_ext_sap=diff_ext_sap_plant;
        diff_ext_so=diff_ext_so_plant;
        prc_sap_so=prc_sap_so_plant;
        prc_ext_sap=prc_ext_sap_plant;
        prc_ext_so=prc_ext_so_plant;
    end;
  run;

  data SvsD_variety_name_global(drop=rc);
  	set SvsD_units;
  	length variety_name_global $60.;
  	if _n_=1 then do;
  	declare hash h_vng(dataset: 'dmproc.pmd_assortment(rename=(GLOBAL_variety_name=variety_name_global))');
  	  rc=h_vng.DefineKey ('variety', 'region');
  	  rc=h_vng.DefineData ('variety_name_global');
  	  rc=h_vng.DefineDone();
  	end;
  	rc=h_vng.find();
  run;

  data SvsD_final(keep= region
                        Product_Line
                        product_line_group
                        order_season
                        season_week_start
                        species
                        series
                        variety
                        process_stage
                        variety_name
                        variety_name_global
                        variety_plc_current
                        variety_plc_future
                        variety_plc_future_date
                        variety_global_plc_current
                        material
                        material_name
                        material_division
                        product_form
                        material_plc_current
                        material_plc_future
                        material_plc_future_date
                        item_measurement
                        sub_unit
                        inv_orders
                        ooh_orders
                        total_orders
                        s_1_cw_history
                        extrapolation
                        adj_extrapolation
                        extrapolated_orders
                        s_1_eos_history
                        s_2_eos_history
                        s_3_eos_history
                        sap_normal_s0
                        sap_constrained_s0
                        unconsumed_demand
                        so_normal_s0
                        diff_ext_sap
                        prc_ext_sap
                        diff_ext_so
                        prc_ext_so
                        diff_sap_so
                        prc_sap_so
                        sap_normal_s1
                        sap_constrained_s1
                        sap_normal_s2
                        sap_constrained_s2
                        sap_normal_s3
                        sap_constrained_s3
                        so_normal_s1
                        so_normal_s2
                        so_normal_s3

                        );
    retain  region
            Product_Line
            product_line_group
            order_season
            season_week_start
            species
            series
            variety
            process_stage
            variety_name
            variety_plc_current
            variety_plc_future
            variety_plc_future_date
            variety_name_global
            variety_global_plc_current
            material
            material_name
            material_division
            product_form
            material_plc_current
            material_plc_future
            material_plc_future_date
            item_measurement
            sub_unit
            inv_orders
            ooh_orders
            total_orders
            s_1_cw_history
            extrapolation
            adj_extrapolation
            extrapolated_orders
            s_1_eos_history
            s_2_eos_history
            s_3_eos_history
            sap_normal_s0
            sap_constrained_s0
            unconsumed_demand
            so_normal_s0
            diff_ext_sap
            prc_ext_sap
            diff_ext_so
            prc_ext_so
            diff_sap_so
            prc_sap_so
            sap_normal_s1
            sap_constrained_s1
            sap_normal_s2
            sap_constrained_s2
            sap_normal_s3
            sap_constrained_s3
            so_normal_s1
            so_normal_s2
            so_normal_s3
;
    set SvsD_variety_name_global;
  run;


  data _null_;
    svsd_report_file=catx('_', compress(put(today(),yymmdd10.),,'kd'), compress(put(time(), time8.),,'kd'));
    call symput('svsd_report_file', strip(svsd_report_file));
    root_report_folder="&svsd_report_folder.";
    date_report_folder="&report_year.&report_week.";
    rc=dcreate(date_report_folder, root_report_folder);
  run;

  x "del &svsd_report_folder.\&report_year.&report_week.\SvdD_report_&svsd_report_file..xlsx"; 

  proc export 
    data=svsd_final 
    dbms=xlsx 
    outfile="&svsd_report_folder.\&report_year.&report_week.\SvdD_report_&svsd_report_file..xlsx" replace label;
    sheet="Created_on_&svsd_report_file.";
  run;

  proc export 
    data=svsd_report_md_raw1
    dbms=xlsx 
    outfile="&svsd_report_folder.\&report_year.&report_week.\SvdD_report_&svsd_report_file..xlsx";
    sheet="Variant";
  run;

  %cleanup_xlsx_bak_folder(cleanup_folder=%str(&svsd_report_folder.\&report_year.&report_week.\));
  
%mend REPORT_SALES_VS_DEMAND;

%REPORT_SALES_VS_DEMAND();