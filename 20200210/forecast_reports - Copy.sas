/***********************************************************************/
/*Type: Report*/
/*Use: Fill in parameters in metadata.xlsx, sheet=Forecast_report and press run*/
/*Purpose: Create forecast report with 4 seperate steps*/
/*OUT: dmfcst1.&region._&product_line_group._&ka_material_division._&season._&current_year_week._&_seasonality.*/
/*     dmfcst2.&region._&product_line_group._&ka_material_division._&season._&current_year_week._&_seasonality.*/
/*     dmfcst3.&region._&product_line_group._&ka_material_division._&season._&current_year_week._&_seasonality.*/
/*     dmfcst4.&region._&product_line_group._&ka_material_division._&season._&current_year_week._&_seasonality.*/
/*     Muptiple excel report files in in upload_report_folder (check configuration.sas)*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\cleanup_xlsx_bak_folder.sas";
%include "&sas_applications_folder.\read_sales_history_weekly.sas";
%include "&sas_applications_folder.\xls_capacity.sas";
%include "&sas_applications_folder.\xls_supply.sas";
%include "&sas_applications_folder.\concatenate_forecast_sm_feedback.sas";
%include "&sas_applications_folder.\read_metadata.sas";

%macro forecast_report_step1(Region=, Product_line_group=, material_division=, seasonality=, Season=, current_year_week=, Supply_vertical_file=, Supply_horizontal_file=, Capacity_vertical_file=, Capacity_horizontal_file=, previous_forecast4_sas_table=);

  %let _seasonality=%sysfunc(tranwrd(%quote(&seasonality.),%str(-),%str(_)));
  %let kda_material_division=%sysfunc(compress(%quote(&material_division.),,kda));
  %let excel_report_fname=&region._&product_line_group._&kda_material_division._&season._&current_year_week._&_seasonality.;
  %let ka_material_division=%sysfunc(compress(%quote(&material_division.),,ka));
  %let sas_report_fname=&region._&product_line_group._&ka_material_division._&season._&current_year_week._&_seasonality.;

  proc sql;
  create table orders_s_aggr_c as
    select region, variety, territory, country, order_season, sum(historical_sales) as historical_sales, sum(actual_sales) as actual_sales 
      from dmproc.orders_all
      where ^missing(order_season) and ^missing(actual_sales) and mat_div in (&material_division.) and product_line_group="&product_line_group." and region="&region." 
      group by region, variety, territory, country, order_season;
  quit;

/*  proc sql;*/
/*  create table orders_s_aggr_t as*/
/*    select region, variety, territory, territory as country, order_season, sum(historical_sales) as historical_sales, sum(actual_sales) as actual_sales */
/*      from dmproc.orders_all*/
/*      where ^missing(order_season) and ^missing(actual_sales) and mat_div in (&material_division.)*/
/*      group by region, variety, territory, order_season;*/
/*  quit;*/

  proc sql;
  create table orders_s_aggr_r as
    select region, variety, region as territory, region as country, order_season, sum(historical_sales) as historical_sales, sum(actual_sales) as actual_sales 
      from dmproc.orders_all
      where ^missing(order_season) and ^missing(actual_sales) and mat_div in (&material_division.) and product_line_group="&product_line_group." and region="&region."
      group by region, variety, order_season;
  quit;

  data orders_s_aggr;
    set orders_s_aggr_c 
        /*orders_s_aggr_t*/ 
        orders_s_aggr_r;
  run;

  data fr;
    set DMPROC.PMD_Assortment;
    length seasons $5.;
    seasons=strip(strip(put(season_week_start, 2.))||'-'||strip(put(season_week_end, 2.)));
    if region="&region.";
  run;

  data fr1;
    set fr;
      if strip(product_line_group) = "&Product_line_group." and current_plc in ('E0', 'E1', 'E2', 'F0', 'F1', 'F2', 'F3', 'F4', 'G0', 'G2') and strip(seasons)="&seasonality." then output;
  run;

  proc sql;
    create table regions4report as 
      select distinct 1 as order, 'Country' as type, territory, country from (select region, territory, country from dmimport.country_lookup union select region, territory, country from dmimport.soldto_nr_lookup) where region="&region." and region^=country /*exception for BI/BI/BI, will also work after split into regions and countries*/
      /*union select distinct 2 as order, 'Territory' as type, territory, territory as country from dmimport.country_lookup where region="&region." and region^=territory /*exception for BI/BI/BI, will also work after split into regions and countries*/
      union select distinct 3 as order, region as type, region as territory, region as country from (select region, territory, country from dmimport.country_lookup union select region, territory, country from dmimport.soldto_nr_lookup) where region="&region.";
  quit;

  proc sql;
    create table allvarieties4report as 
    select order, type, variety, territory, country from regions4report
    join (select distinct variety from fr1) on 1=1
    order by variety, order, territory, country;
  quit;

  proc sql;
    create table fr2 as
    select 
      order,
      region,
      type, 
      seasons, 
      territory, 
      country, 
      product_line,
      species,
      series,
      a.variety,
      variety_name,
      current_plc as plc,
      future_plc,
      future_plc_active_date as valid_from_date,
      global_current_plc as global_plc,
      global_future_plc,
      global_future_plc_active_date as global_valid_from_date,
      hash_product_line,
      hash_species_name,
      season_week_start
    from allvarieties4report a
    left join fr1 f on f.variety=a.variety;
  quit;

  %let previous_season=%eval(&season.-1);

  %let current_year=%substr(&current_year_week., 1, 4);
  %let current_week=%substr(&current_year_week., 5, 2);
  %let season_start_week=%scan(&seasonality., 1, "-");
  %let season_end_week=%scan(&seasonality., 2, "-");
  %let last_season_year=%eval(&current_year.-1);
  %let last_season_week=&current_week.;

  %if &season_end_week >=52 %then %do;
    %let end_season_year=%eval(&season.-1);
  %end; %else %do;
    %let end_season_year=&season.;
  %end;
  %let end_season_week=&season_end_week.;

  %put RPS: &=previous_season &=last_season_year &=last_season_week &=end_season_year, &=end_season_week, &=previous_season;

  %read_sales_history_weekly(last_season_year=&last_season_year., 
                            last_season_week=&last_season_week., 
                            end_season_year=&end_season_year., 
                            end_season_week=&end_season_week., 
                            previous_season=&previous_season., 
                            region=&region., 
                            material_division=%quote(&material_division.), 
                            product_line_group=&product_line_group.);

  data fr3(drop=rc historical_sales actual_sales order_season);
    length rc 8.;
    set fr2;
    length historical_sales actual_sales s1 s2 s3 order_season ytd percentage extrapolation prev_demand1 prev_demand2 prev_demand3 8.;
    if _n_=1 then do;
      declare hash horders(dataset: 'orders_s_aggr');
        rc=horders.DefineKey ('region', 'variety', 'order_season', 'territory', 'country');
        rc=horders.DefineData ('historical_sales');
        rc=horders.DefineDone();
      declare hash aorders(dataset: 'orders_s_aggr');
        rc=aorders.DefineKey ('region', 'variety', 'order_season', 'territory', 'country');
        rc=aorders.DefineData ('actual_sales');
        rc=aorders.DefineDone();
      declare hash sales_percentage(dataset: 'sales_percentage');
        rc=sales_percentage.DefineKey ('hash_species_name');
        rc=sales_percentage.DefineData ('percentage');
        rc=sales_percentage.DefineDone();
    end;

    percentage=coalesce(percentage,0);
    order_season=&season0.;
    rc=aorders.find(); 
    ytd=coalesce(actual_sales,0);

    order_season=&season1.;
    rc=horders.find();
    s1=coalesce(historical_sales,0);

    call missing(historical_sales);
    order_season=&season2.;
    rc=horders.find();
    s2=coalesce(historical_sales,0);

    call missing(historical_sales);
    order_season=&season3.;
    rc=horders.find();
    s3=coalesce(historical_sales,0);

    rc=sales_percentage.find();

    if percentage^=0 then do;
      extrapolation=coalesce(round(ytd/percentage, 1),0);
    end; else do;
      extrapolation=0;
    end;
  run;

  proc sql noprint;
  create table fr4 as
    select   a.*,
        case when ^missing(b.ytd) then a.ytd/b.ytd 
        end as current_season_split,
        case when ^missing(b.s1) then a.s1/b.s1
        end as last_season_split
    from fr3 a 
    left join fr3 b on a.variety=b.variety and b.country="&region."
    order by variety, order, territory, country;
  quit;

  proc sql;
    create table fr4_series as
    select   a.order,
            a.series,
            a.country,
            case when ^missing(b.ytd) then a.ytd/b.ytd 
            end as current_season_series_split,
            case when ^missing(b.s1) then a.s1/b.s1
             end as last_season_series_split 
    from (select order, series, country, sum(ytd) as ytd, sum(s1) as s1 from fr3 group by order, series, country) a
    left join (select series, sum(ytd) as ytd, sum(s1) as s1 from fr3 where country="&region." group by series) b on a.series=b.series;
  quit;

  proc sql;
    create table fr4_species as
    select   a.order,
            a.species,
            a.country,
            case when ^missing(b.ytd) then a.ytd/b.ytd 
            end as current_season_species_split,
            case when ^missing(b.s1) then a.s1/b.s1
             end as last_season_species_split 
    from (select order, species, country, sum(ytd) as ytd, sum(s1) as s1 from fr3 group by order, species, country) a
    left join (select species, sum(ytd) as ytd, sum(s1) as s1 from fr3 where country="&region." group by species) b on a.species=b.species;
  quit;

  proc sql;
    create table fr4_general as
    select   a.order,
            a.country,
            case when ^missing(b.ytd) then a.ytd/b.ytd
            end as current_season_general_split,
            case when ^missing(b.s1) then a.s1/b.s1
             end as last_season_general_split 
    from (select order, country, sum(ytd) as ytd, sum(s1) as s1 from fr3 group by order, country) a
    left join (select sum(ytd) as ytd, sum(s1) as s1 from fr3 where country="&region.") b on 1=1;
  quit;

  proc sql;
    create table fr5 as
    select   a.*, 
            coalesce(b.current_season_series_split,0) as current_season_series_split, 
            coalesce(b.last_season_series_split,0) as last_season_series_split, 
            coalesce(c.current_season_species_split,0) as current_season_species_split, 
            coalesce(c.last_season_species_split,0) as last_season_species_split, 
            coalesce(d.current_season_general_split,0) as current_season_general_split, 
            coalesce(d.last_season_general_split,0) as last_season_general_split,
            case when percentage >= 0.8 
            then 
              coalesce(a.current_season_split, b.current_season_series_split, c.current_season_species_split, d.current_season_general_split)
            else
              coalesce(a.last_season_split, b.last_season_series_split, c.last_season_species_split, d.last_season_general_split) 
            end as country_split
    from fr4 a
    left join fr4_series b on a.order=b.order and a.series=b.series and a.country=b.country
    left join fr4_species c on a.order=c.order and a.species=c.species and a.country=c.country
    left join fr4_general d on a.order=d.order and a.country=d.country
    order by variety, order, territory, country;
  quit;

  data fr6;
    set fr5;
    current_season_split=coalesce(current_season_split,0);
    last_season_split=coalesce(last_season_split,0);
  run;

  data fr7(drop=rc);
  length rc 8.;
    set fr6;
  %if "&previous_forecast4_sas_table."^="" %then %do;
    if _n_=1 then do;
/*REVIEW ROUND - SAME SEASON*/
      declare hash prvdmnd (dataset: "dmfcst4.&previous_forecast4_sas_table.(DROP= prev_demand1 prev_demand2 prev_demand3 rename=(so_demand1=prev_demand1 so_demand2=prev_demand2 so_demand3=prev_demand3))");
/*NEW ROUND - NEXT SEASON*/
/*      declare hash prvdmnd (dataset: "dmfcst4.&previous_forecast4_sas_table.(DROP= prev_demand1 prev_demand2 prev_demand3 rename=(so_demand1=prev_demand2 so_demand2=prev_demand3 so_demand3=prev_demand3))");*/

        rc=prvdmnd.DefineKey ('region', 'variety', 'country');
        rc=prvdmnd.DefineData ('prev_demand1', 'prev_demand2', 'prev_demand3');
        rc=prvdmnd.DefineDone();
    end;
    rc=prvdmnd.find();
    prev_demand1=coalesce(prev_demand1,0);
    prev_demand2=coalesce(prev_demand2,0);
    prev_demand3=coalesce(prev_demand3,0);
  %end;
  run;

  %if %sysfunc(find(&product_line_group.,Cut)) ge 0 %then %let is_cutting=Y;
    %else %let is_cutting=N;

  data fr8;
    set fr7;
    length pm_demand1 pm_demand2 pm_demand3 8.;
    pm_demand1=prev_demand1;
    pm_demand2=prev_demand2;
    pm_demand3=prev_demand3;

    if country^="&region." then do;
      pm_demand1=0; 
      pm_demand2=0; 
      pm_demand3=0;
    end; 
/* *********************************************************************** 40K CHECK ******************************************/
/*     %if "&is_cutting."="Y" %then %do;*/
/*      if country="FPS" and plc='E2' then do;*/
/*        if missing(pm_demand1) or 0<pm_demand1<40000 then pm_demand1=40000;*/
/*        if missing(pm_demand2) or 0<pm_demand2<40000 then pm_demand2=40000;*/
/*        if missing(pm_demand3) or 0<pm_demand3<40000 then pm_demand3=40000;*/
/*      end;*/
/*      if country="BI" and plc='E2' then do;*/
/*        if missing(pm_demand1) or 0<pm_demand1<10000 then pm_demand1=10000;*/
/*        if missing(pm_demand2) or 0<pm_demand2<10000 then pm_demand2=10000;*/
/*        if missing(pm_demand3) or 0<pm_demand3<10000 then pm_demand3=10000;*/
/*      end;*/
/*    %end;*/
/*  run;*/
/* *********************************************************************** 40K CHECK ******************************************/
  data fr9 (drop=next_season: first_day_of_next_season:);
    set fr8;
    format valid_from_date first_day_of_next_season1 first_day_of_next_season2 first_day_of_next_season3 yymmdd10.;
    next_season1=&season.+1;
    next_season2=&season.+2;
    next_season3=&season.+3;
    first_day_of_next_season1=input(put(next_season1, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
    first_day_of_next_season2=input(put(next_season2, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
    first_day_of_next_season3=input(put(next_season3, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
    if plc='G2' or (future_plc='G2' and valid_from_date <= first_day_of_next_season1) then do;
      call missing(pm_demand1);
    end;
    if plc='G2' or (future_plc='G2' and valid_from_date <= first_day_of_next_season2) then do;
      call missing(pm_demand2);
    end;    
    if plc='G2' or (future_plc='G2' and valid_from_date <= first_day_of_next_season3) then do;
      call missing(pm_demand3);
    end;
    pm_demand1=coalesce(pm_demand1,0);
    pm_demand2=coalesce(pm_demand2,0);
    pm_demand3=coalesce(pm_demand3,0);
  run;

  proc sql noprint;
    select strip(dir_lvl31) into :supply_and_capacity_folder trimmed from forecast_report_folders;
  quit;

  %if "&supply_horizontal_file."^="" or "&supply_vertical_file."^="" %then %do;
    %if "&supply_horizontal_file."^="" %then %do;
      x "copy ""&supply_horizontal_file."" ""&supply_and_capacity_folder.\""";
      %read_supply_horizontal(horizontal_supply_file=%quote(&supply_horizontal_file.));
    %end;
    %if "&supply_vertical_file."^="" %then %do;
      x "copy ""&supply_vertical_file."" ""&supply_and_capacity_folder.\""";
      %read_supply_vertical(vertical_supply_file=%quote(&supply_vertical_file.));
    %end;
    proc sql;
      create table fr10 as
        select a.*, coalesce(b.supply, 0) as supply from fr9 a
        left join (select variety, sum(supply) as supply from dmimport.supply group by variety) b on a.variety=b.variety and a.country="&region.";
    quit;
  %end; %else %do;
    data fr10;
      set fr9;
      length supply 8.;
      supply=0;
      run;
  %end;

  %if "&capacity_horizontal_file."^="" or "&capacity_vertical_file."^="" %then %do;
    %if "&capacity_horizontal_file."^="" %then %do;
      x "copy ""&capacity_horizontal_file."" ""&supply_and_capacity_folder.\""";
      %read_capacity_horizontal(horizontal_capacity_file=%quote(&capacity_horizontal_file.));
    %end;
    %if "&capacity_vertical_file."^="" %then %do;
      x "copy ""&capacity_vertical_file."" ""&supply_and_capacity_folder.\""";
      %read_capacity_vertical(vertical_capacity_file=%quote(&capacity_vertical_file.));
    %end;
    proc sql;
      create table fr11 as
        select a.*, coalesce(b.capacity, 0) as capacity from fr10 a
        left join (select variety, sum(capacity) as capacity from dmimport.capacity group by variety) b on a.variety=b.variety and a.country="&region.";
    quit;
  %end; %else %do;
    data fr11;
      set fr10;
      length capacity 8.;
      capacity=0;
      run;
  %end;

  data fr_end;
    set fr11;
  run;

  proc sort data=fr_end;
    by product_line species series variety order territory country;
  run;

  %let FR_COLUMNS_FOR_TEMPLATE=order seasons Territory Country Product_Line Species Series variety Variety_name 
       plc Future_PLC valid_from_date 
       global_plc Global_Future_PLC global_valid_from_date empty_column1 
       s3 s2 s1 empty_column2 
       ytd percentage extrapolation empty_column3 
       country_split supply capacity remark empty_column4 
       pm_demand1 sm_demand1 prev_demand1 assumption1 
       pm_demand2 sm_demand2 prev_demand2 assumption2 
       pm_demand3 sm_demand3 prev_demand3 assumption3;

  data FOR_TEMPLATE(keep=&FR_COLUMNS_FOR_TEMPLATE.);
    retain &FR_COLUMNS_FOR_TEMPLATE.;
    length sm_demand1 sm_demand2 sm_demand3 8.;
    length assumption1 assumption2 assumption3 empty_column1 empty_column2 empty_column3 empty_column4 remark $1.;
    set fr_end;
    percentage=percentage; 
    country_split=country_split; 
    pm_demand1=round(pm_demand1,1);
    pm_demand2=round(pm_demand2,1);
    pm_demand3=round(pm_demand3,1);
    prev_demand1=round(prev_demand1,1);
    prev_demand2=round(prev_demand2,1);
    prev_demand3=round(prev_demand3,1);
    sm_demand1=round(sm_demand1,1);
    sm_demand2=round(sm_demand2,1);
    sm_demand3=round(sm_demand3,1);
    s3=round(s3,1);
    s2=round(s2,1);
    s1=round(s1,1);
    ytd=round(ytd,1);
    extrapolation=round(extrapolation,1);
  run;

  proc sort data=FOR_TEMPLATE;
    by product_line species series variety order territory country;
  run; 

  data dmfcst1.&sas_report_fname.;
    set fr_end;
  run;

  proc sql noprint;
    select dir_lvl32 into :pm_folder trimmed from forecast_report_folders;
  quit;

  proc export 
    data=fr_end 
    dbms=xlsx 
    outfile="&pm_folder.\&excel_report_fname..xlsx" 
    replace;
    SHEET="Forecast_step1"; 
  run;

  proc export 
    data=FOR_TEMPLATE 
    dbms=xlsx 
    outfile="&pm_folder.\&excel_report_fname._FOR_TEMPLATE.xlsx" 
    replace;
    SHEET="Variety level fcst"; 
  run;

  %cleanup_xlsx_bak_folder(cleanup_folder=%str(&pm_folder.));

%mend forecast_report_step1;

%macro forecast_report_step2(  Region=, Product_line_group=, material_division=, seasonality=, Season=, current_year_week=, pm_feedback=);

  %let _seasonality=%sysfunc(tranwrd(%quote(&seasonality.),%str(-),%str(_)));
  %let kda_material_division=%sysfunc(compress(%quote(&material_division.),,kda));
  %let excel_report_fname=&region._&product_line_group._&kda_material_division._&season._&current_year_week._&_seasonality.;
  %let ka_material_division=%sysfunc(compress(%quote(&material_division.),,ka));
  %let sas_report_fname=&region._&product_line_group._&ka_material_division._&season._&current_year_week._&_seasonality.;

  PROC IMPORT OUT=pm_feedback_raw 
              DATAFILE="&pm_feedback."
              DBMS=  EXCELCS  REPLACE;
              SHEET="Variety level fcst"; 
  RUN;

  proc contents data=pm_feedback_raw out=pm_feedback_contents noprint;
  run;

  data pm_feedback_cols;
    length varnum 8. columnname $32.;
    varnum=4; /*Excel column D*/
    columnname="Country"; 
    output;
    varnum=8; /*Excel column H*/
    columnname="Variety";
    output;
    varnum=10; /*Excel column J*/
    columnname="PLC";
    output;
    varnum=30; /*Excel column AD*/
    columnname="pmf_demand1";
    output;
    varnum=34; /*Excel column AH*/
    columnname="pmf_demand2";
    output;
    varnum=38; /*Excel column AL*/
    columnname="pmf_demand3";
    output;
    varnum=33; /*Excel column AG*/
    columnname="pmf_assm1";
    output;
    varnum=37; /*Excel column AK*/
    columnname="pmf_assm2";
    output;
    varnum=41; /*Excel column AO*/
    columnname="pmf_assm3";
    output;
  run;

  proc sql noprint;
    select compress(name||'=_'||columnname) into :renamestring separated by ' ' from pm_feedback_cols fcols
    left join pm_feedback_contents frcst on fcols.varnum=frcst.varnum;
  quit;

  data pm_feedback1 (keep=variety country PLC sum_of_3_seasons pmf_demand1 pmf_demand2 pmf_demand3 pmf_assm1 pmf_assm2 pmf_assm3);
    set pm_feedback_raw(rename=(&renamestring.));
    length country $6. PLC $2. variety pmf_demand1 pmf_demand2 pmf_demand3 8.;
    length pmf_assm1 pmf_assm2 pmf_assm3 $1000.;
    country=strip(_country);
    PLC=strip(_PLC);
    variety=input(strip(_variety), 8.);
    pmf_demand1=input(_pmf_demand1, comma20.);
    pmf_demand2=input(_pmf_demand2, comma20.);
    pmf_demand3=input(_pmf_demand3, comma20.);
    if missing(pmf_demand1) then pmf_demand1=0;
    if missing(pmf_demand2) then pmf_demand2=0;
    if missing(pmf_demand3) then pmf_demand3=0;
    pmf_assm1=strip(_pmf_assm1);
    pmf_assm2=strip(_pmf_assm2);
    pmf_assm3=strip(_pmf_assm3);
    sum_of_3_seasons=pmf_demand1+pmf_demand2+pmf_demand3;
    if ^missing(Variety) and country="&region." then output;
  run;

  %if %sysfunc(find(&product_line_group.,Cut)) ge 0 %then %let is_cutting=Y;
    %else %let is_cutting=N;

  data pm_feedback2;
    set pm_feedback1;
/* *********************************************************************** 40K CHECK ******************************************/
/*    %if "&is_cutting."="Y" %then %do;*/
/*      if country="FPS" and plc='E2' then do;*/
/*        if 0<pmf_demand1<40000 then pmf_demand1=40000;*/
/*        if 0<pmf_demand2<40000 then pmf_demand2=40000;*/
/*        if 0<pmf_demand3<40000 then pmf_demand3=40000;*/
/*      end;*/
/*      if country="BI" and plc='E2' then do;*/
/*        if 0<pmf_demand1<10000 then pmf_demand1=10000;*/
/*        if 0<pmf_demand2<10000 then pmf_demand2=10000;*/
/*        if 0<pmf_demand3<10000 then pmf_demand3=10000;*/
/*      end;*/
/*    %end;*/
/* *********************************************************************** 40K CHECK ******************************************/
  run;

  proc sql;
    create table pmf1 as
    select a.*, 
          b.pmf_demand1,
          b.pmf_demand2,
          b.pmf_demand3,
          pmf_assm1,
          pmf_assm2,
          pmf_assm3,
          round((b.pmf_demand1*a.country_split),1) as pmf_split_demand1, 
          round((b.pmf_demand2*a.country_split),1) as pmf_split_demand2, 
          round((b.pmf_demand3*a.country_split),1) as pmf_split_demand3 
    from dmfcst1.&sas_report_fname. a 
    left join pm_feedback2 b on a.variety=b.variety;
  quit;

  proc sql;
    create table pmf2 as
    select a.*, round(b.supply_total*a.country_split) as pmf_supply from pmf1 a
    left join (select variety, supply as supply_total from pmf1 where country="&region.") b on a.variety=b.variety;
  quit;

  proc sql;
    create table pmf3 as
    select a.*, round(b.capacity_total*a.country_split) as pmf_capacity from pmf2 a
    left join (select variety, capacity as capacity_total from pmf1 where country="&region.") b on a.variety=b.variety;
  quit;

  data pmf4 (drop=next_season: first_day_of_next_season:);
    set pmf3;
    format valid_from_date first_day_of_next_season1 first_day_of_next_season2 first_day_of_next_season3 yymmdd10.;
    next_season1=&season.+1;
    next_season2=&season.+2;
    next_season3=&season.+3;
    first_day_of_next_season1=input(put(next_season1, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
    first_day_of_next_season2=input(put(next_season2, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
    first_day_of_next_season3=input(put(next_season3, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
    if plc='G2' or (future_plc='G2' and valid_from_date <= first_day_of_next_season1) then do;
      call missing(pmf_split_demand1);
    end;
    if plc='G2' or (future_plc='G2' and valid_from_date <= first_day_of_next_season2) then do;
      call missing(pmf_split_demand2);
    end;    
    if plc='G2' or (future_plc='G2' and valid_from_date <= first_day_of_next_season3) then do;
      call missing(pmf_split_demand3);
    end;
    pmf_split_demand1=coalesce(pmf_split_demand1,0);
    pmf_split_demand2=coalesce(pmf_split_demand2,0);
    pmf_split_demand3=coalesce(pmf_split_demand3,0);
  run;

  data pmf5;
    set pmf4;
    length pmf_sm_demand1 pmf_sm_demand2 pmf_sm_demand3 8.;
    pmf_sm_demand1=pmf_split_demand1;
    pmf_sm_demand2=pmf_split_demand2;
    pmf_sm_demand3=pmf_split_demand3;
  run;

  data pmf_end;
    set pmf5;
  run;

  proc sort data=pmf_end;
    by product_line species series variety order territory country;
  run; 

  data dmfcst2.&sas_report_fname.;
    set pmf_end;
  run;

  %let PMF_COLUMNS_FOR_TEMPLATE=order seasons Territory Country Product_Line Species Series variety Variety_name 
     plc Future_PLC valid_from_date 
     global_plc Global_Future_PLC global_valid_from_date empty_column1 
     s3 s2 s1 empty_column2 
     ytd percentage extrapolation empty_column3 
     country_split pmf_supply pmf_capacity remark empty_column4 
     pmf_split_demand1 pmf_sm_demand1 prev_demand1 pmf_assm1 
     pmf_split_demand2 pmf_sm_demand2 prev_demand2 pmf_assm2 
     pmf_split_demand3 pmf_sm_demand3 prev_demand3 pmf_assm3;

  data FOR_TEMPLATE_COUNTRIES(keep=&PMF_COLUMNS_FOR_TEMPLATE.);
    retain &PMF_COLUMNS_FOR_TEMPLATE.;
    length empty_column1 empty_column2 empty_column3 empty_column4 remark $1.;
    set pmf_end;
    percentage=percentage; 
    country_split=country_split; 
    pmf_split_demand1=round(pmf_split_demand1,1);
    pmf_split_demand2=round(pmf_split_demand2,1);
    pmf_split_demand3=round(pmf_split_demand3,1);
    prev_demand1=round(prev_demand1,1);
    prev_demand2=round(prev_demand2,1);
    prev_demand3=round(prev_demand3,1);
    pmf_sm_demand1=round(pmf_sm_demand1,1);
    pmf_sm_demand2=round(pmf_sm_demand2,1);
    pmf_sm_demand3=round(pmf_sm_demand3,1);
    s3=round(s3,1);
    s2=round(s2,1);
    s1=round(s1,1);
    ytd=round(ytd,1);
    extrapolation=round(extrapolation,1);
  run;

  proc sort data=FOR_TEMPLATE_COUNTRIES;
    by product_line species series variety order territory country;
  run; 

  proc sql noprint;
    select dir_lvl34 into :sm_folder trimmed from forecast_report_folders;
  quit;

  proc export 
    data=pmf_end 
    dbms=xlsx 
    outfile="&sm_folder.\&excel_report_fname..xlsx" 
    replace;
    SHEET="Forecast_step2"; 
  run;

  proc export 
    data=FOR_TEMPLATE_COUNTRIES 
    dbms=xlsx 
    outfile="&sm_folder.\&excel_report_fname._ALL.xlsx" 
    replace;
    SHEET="Forecast_step2"; 
  run;

  proc sql noprint;
    select count(distinct(country)) into :country_cnt from FOR_TEMPLATE_COUNTRIES where country^="&region.";
    select distinct(country) into :country_list separated by '#' from FOR_TEMPLATE_COUNTRIES where country^="&region.";
  quit;

  %do ci=1 %to &country_cnt.;
    %let country=%scan(&country_list., &ci., '#');
    data pmf_country;
      set FOR_TEMPLATE_COUNTRIES;
      if country="&country." then output;
    run;

    proc export 
      data=pmf_country 
      dbms=xlsx 
      outfile="&sm_folder.\&excel_report_fname._&country..xlsx" 
      replace;
      SHEET="Variety level fcst";
    run;
  %end;

  %cleanup_xlsx_bak_folder(cleanup_folder=%str(&sm_folder.));

%mend forecast_report_step2;

%macro forecast_report_step3(  Region=, Product_line_group=, material_division=, seasonality=, Season=, current_year_week=, sm_feedback_folder=);

  %let _seasonality=%sysfunc(tranwrd(%quote(&seasonality.),%str(-),%str(_)));
  %let kda_material_division=%sysfunc(compress(%quote(&material_division.),,kda));
  %let excel_report_fname=&region._&product_line_group._&kda_material_division._&season._&current_year_week._&_seasonality.;
  %let ka_material_division=%sysfunc(compress(%quote(&material_division.),,ka));
  %let sas_report_fname=&region._&product_line_group._&ka_material_division._&season._&current_year_week._&_seasonality.;

  %concatenate_forecast_sm_feedback(sm_feedback_folder=%quote(&sm_feedback_folder.));

  data smf1;
    set dmfcst2.&sas_report_fname.;
  run;

  proc sql;
    create table forecast_sm_feedback_demand1 as
    select * from forecast_sm_feedback_demand 
    union (select "&region." as country, 
                  variety, 
                  sum(smf_demand1) as smf_demand1, 
                  sum(smf_demand2) as smf_demand2, 
                  sum(smf_demand3) as smf_demand3 
          from forecast_sm_feedback_demand 
          group by variety)
    order by variety, country;
  quit;

  proc sql;
    create table forecast_sm_feedback_demand2 as
    select a.*, b.smf_total_demand1, b.smf_total_demand2, b.smf_total_demand3 
    from forecast_sm_feedback_demand1 a 
    left join (select variety, 
                      sum(smf_demand1) as smf_total_demand1, 
                      sum(smf_demand2) as smf_total_demand2, 
                      sum(smf_demand3) as smf_total_demand3 
          from forecast_sm_feedback_demand 
          group by variety) b on a.variety=b.variety
    order by variety, country;
  quit;

  proc sql;
    create table smf2 as
    select a.*, b.smf_demand1 as smf_demand1_raw, b.smf_demand2 as smf_demand2_raw, b.smf_demand3 as smf_demand3_raw, 
                b.smf_demand1, b.smf_demand2, b.smf_demand3, 
                b.smf_total_demand1, b.smf_total_demand2, b.smf_total_demand3, c.smf_assm1, c.smf_assm2, c.smf_assm3 from smf1 a
    left join forecast_sm_feedback_demand2 b on a.variety=b.variety and a.country=b.country
    left join forecast_sm_feedback_assm c on a.variety=c.variety and a.country=c.country;
  quit;

  data smf3;
    set smf2;
    if country="&region." then do;
      smf_assm1=pmf_assm1;
      smf_assm2=pmf_assm2;
      smf_assm3=pmf_assm3;
    end;
  run;

  %if %sysfunc(find(&product_line_group.,Cut)) ge 0 %then %let is_cutting=Y;
  %else %let is_cutting=N;

  data smf4;
    set smf3;
  
/* *********************************************************************** 40K CHECK ******************************************/
/*    %if "&is_cutting."="Y" %then %do;*/
/*      if region="FPS" and plc='E2' then do;*/
/*        if 0<smf_total_demand1<40000 then smf_demand1=smf_demand1/smf_total_demand1*40000;*/
/*        if 0<smf_total_demand2<40000 then smf_demand2=smf_demand2/smf_total_demand2*40000;*/
/*        if 0<smf_total_demand3<40000 then smf_demand3=smf_demand3/smf_total_demand3*40000;*/
/*      end;*/
/*      if region="BI" and plc='E2' then do;*/
/*        if 0<smf_total_demand1<10000 then smf_demand1=smf_demand1/smf_total_demand1*40000;*/
/*        if 0<smf_total_demand2<10000 then smf_demand1=smf_demand2/smf_total_demand2*40000;*/
/*        if 0<smf_total_demand3<10000 then smf_demand1=smf_demand3/smf_total_demand3*40000;*/
/*      end;*/
/*    %end;*/
/* *********************************************************************** 40K CHECK ******************************************/
    
  run;

  data smf5 (drop=next_season: first_day_of_next_season:);
    set smf4;
    format valid_from_date first_day_of_next_season1 first_day_of_next_season2 first_day_of_next_season3 yymmdd10.;
    next_season1=&season.+1;
    next_season2=&season.+2;
    next_season3=&season.+3;
    first_day_of_next_season1=input(put(next_season1, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
    first_day_of_next_season2=input(put(next_season2, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
    first_day_of_next_season3=input(put(next_season3, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
    if plc='G2' or (future_plc='G2' and valid_from_date <= first_day_of_next_season1) then do;
      call missing(smf_demand1);
    end;
    if plc='G2' or (future_plc='G2' and valid_from_date <= first_day_of_next_season2) then do;
      call missing(smf_demand2);
    end;    
    if plc='G2' or (future_plc='G2' and valid_from_date <= first_day_of_next_season3) then do;
      call missing(smf_demand3);
    end;
    smf_demand1=coalesce(round(smf_demand1,1),0);
    smf_demand2=coalesce(round(smf_demand2,1),0);
    smf_demand3=coalesce(round(smf_demand3,1),0);
  run;
  
  data smf_end;
    set smf5;
  run;

  proc sort data=smf_end;
    by product_line species series variety order territory country;
  run; 

  data dmfcst3.&sas_report_fname.;
    set smf_end;
  run;

  %let SMF_COLUMNS_FOR_TEMPLATE=order seasons Territory Country Product_Line Species Series variety Variety_name 
     plc Future_PLC valid_from_date 
     global_plc Global_Future_PLC global_valid_from_date empty_column1 
     s3 s2 s1 empty_column2 
     ytd percentage extrapolation empty_column3 
     country_split pmf_supply pmf_capacity remark empty_column4 
     pmf_split_demand1 smf_demand1 prev_demand1 smf_assm1 
     pmf_split_demand2 smf_demand2 prev_demand2 smf_assm2 
     pmf_split_demand3 smf_demand3 prev_demand3 smf_assm3;

  data SMF_FOR_TEMPLATE(keep=&SMF_COLUMNS_FOR_TEMPLATE.);
    retain &SMF_COLUMNS_FOR_TEMPLATE.;
    length empty_column1 empty_column2 empty_column3 empty_column4 remark $1.;
    set smf_end;
    percentage=percentage; 
    country_split=country_split; 
    pmf_split_demand1=round(pmf_split_demand1,1);
    pmf_split_demand2=round(pmf_split_demand2,1);
    pmf_split_demand3=round(pmf_split_demand3,1);
    prev_demand1=round(prev_demand1,1);
    prev_demand2=round(prev_demand2,1);
    prev_demand3=round(prev_demand3,1);
    smf_demand1=round(smf_demand1,1);
    smf_demand2=round(smf_demand2,1);
    smf_demand3=round(smf_demand3,1);
    s3=round(s3,1);
    s2=round(s2,1);
    s1=round(s1,1);    
    ytd=round(ytd,1);
    extrapolation=round(extrapolation,1);
  run;

  proc sort data=SMF_FOR_TEMPLATE;
    by product_line species series variety order territory country;
  run; 

  proc sql noprint;
    select dir_lvl36 into :consolidated_folder trimmed from forecast_report_folders;
  quit;


  proc export 
    data=smf_end 
    dbms=xlsx 
    outfile="&consolidated_folder.\&excel_report_fname..xlsx" 
    replace;
    SHEET="Forecast_step3"; 
  run;

  proc export 
    data=SMF_FOR_TEMPLATE 
    dbms=xlsx 
    outfile="&consolidated_folder.\&excel_report_fname._FOR_TEMPLATE.xlsx" 
    replace;
    SHEET="Forecast_step3"; 
  run;

  %cleanup_xlsx_bak_folder(cleanup_folder=%str(&consolidated_folder.));

%mend forecast_report_step3;

%macro forecast_report_step4(  Region=, Product_line_group=, material_division=, seasonality=, Season=, current_year_week=, signoff_file=);

  %let _seasonality=%sysfunc(tranwrd(%quote(&seasonality.),%str(-),%str(_)));
  %let kda_material_division=%sysfunc(compress(%quote(&material_division.),,kda));
  %let excel_report_fname=&region._&product_line_group._&kda_material_division._&season._&current_year_week._&_seasonality.;
  %let ka_material_division=%sysfunc(compress(%quote(&material_division.),,ka));
  %let sas_report_fname=&region._&product_line_group._&ka_material_division._&season._&current_year_week._&_seasonality.;

    PROC IMPORT OUT=signoff_file_raw 
              DATAFILE="&signoff_file."
              DBMS=  EXCELCS  REPLACE;
              SHEET="Variety level fcst"; 
  RUN;

  proc contents data=signoff_file_raw out=so_file_contents noprint;
  run;

  data signoff_file_cols;
    length varnum 8. columnname $32.;
    varnum=4; /*Excel column D*/
    columnname="Country"; 
    output;
    varnum=8; /*Excel column H*/
    columnname="Variety";
    output;
    varnum=10; /*Excel column J*/
    columnname="PLC";
    output;
    varnum=31; /*Excel column AE*/
    columnname="so_demand1";
    output;
    varnum=35; /*Excel column AI*/
    columnname="so_demand2";
    output;
    varnum=39; /*Excel column AM*/
    columnname="so_demand3";
    output;
    varnum=33; /*Excel column AG*/
    columnname="so_assm1";
    output;
    varnum=37; /*Excel column AK*/
    columnname="so_assm2";
    output;
    varnum=41; /*Excel column AO*/
    columnname="so_assm3";
    output;
  run;

  proc sql noprint;
    select compress(name||'=_'||columnname) into :renamestring separated by ' ' from signoff_file_cols fcols
    left join so_file_contents frcst on fcols.varnum=frcst.varnum;
  quit;

  data signoff_file1 (keep=variety country PLC so_demand1 so_demand2 so_demand3 so_assm1 so_assm2 so_assm3);
    set signoff_file_raw(rename=(&renamestring.));
    length country $6. PLC $2. variety so_demand1 so_demand2 so_demand3 8.;
    length so_assm1 so_assm2 so_assm3 $1000.;
    country=strip(_country);
    PLC=strip(_PLC);
    variety=input(strip(_variety), 8.);
    so_demand1=input(_so_demand1, comma20.);
    so_demand2=input(_so_demand2, comma20.);
    so_demand3=input(_so_demand3, comma20.);
    so_assm1=strip(_so_assm1);
    so_assm2=strip(_so_assm2);
    so_assm3=strip(_so_assm3);
    if missing(so_demand1) then so_demand1=0;
    if missing(so_demand2) then so_demand2=0;
    if missing(so_demand3) then so_demand3=0;
    if ^missing(Variety) then output;
  run;

  data so1;
    set dmfcst3.&sas_report_fname.;
  run;

  proc sql;
    create table so2 as
    select a.*, b.so_demand1, b.so_demand2, b.so_demand3, b.so_assm1, b.so_assm2, b.so_assm3 
    from so1 a
    left join signoff_file1 b on a.variety=b.variety and a.country=b.country;
  quit;

  data so_end;
    set so2;
  run;

  proc sort data=so_end;
    by product_line species series variety order territory country;
  run; 

  data dmfcst4.&sas_report_fname.;
    set so_end;
  run;

%mend forecast_report_step4;

%macro forecast_reports();
  
  %read_metadata(sheet=Forecast_reports);

  proc sql noprint;
    select count(*) into :report_cnt from dmimport.Forecast_reports_md;
  quit;
  
  %do ii=1 %to &report_cnt.;

    data _null_;
      set dmimport.Forecast_reports_md;
      if _n_=&ii then do;
        call symput('region', strip(region));
        call symput('product_line_group', strip(product_line_group));
        call symput('material_division', '"'||strip(tranwrd(material_division, ',', '", "'))||'"');
        call symput('season', strip(season));
        call symput('current_year_week', strip(current_year_week));
        call symput('seasonality', strip(seasonality));
        call symput('supply_file', strip(supply_file));
        call symput('Supply_vertical_file', strip(Supply_vertical_file));
        call symput('Supply_horizontal_file', strip(Supply_horizontal_file));
        call symput('Capacity_vertical_file', strip(Capacity_vertical_file));
        call symput('Capacity_horizontal_file', strip(Capacity_horizontal_file));
        call symput('mat_div', strip(mat_div));
        call symput('previous_forecast4_sas_table', strip(previous_forecast4_sas_table));
        call symput('step1', strip(step1));
        call symput('step2', strip(step2));
        call symput('step3', strip(step3));
        call symput('step4', strip(step4));
        call symput('pm_feedback', strip(pm_feedback));
        call symput('sm_feedback_folder', strip(sm_feedback_folder));
        call symput('signoff_file', strip(signoff_file));
      end;
    run;

    %if "&step1."="Y" or "&step2."="Y" or "&step3."="Y" or "&step4."="Y" %then %do;
      %let season0=&season.;
      %let season1=%eval(&season.-1);
      %let season2=%eval(&season.-2);
      %let season3=%eval(&season.-3);
      %let _material_division=%sysfunc(compress(%quote(&material_division.),,kda));

      %put Report_variables &region. &product_line_group. &material_division. &product_line_group. &season. &current_year_week. &seasonality. &mat_div. &previous_forecast4_sas_table. &step1.;
    
      data forecast_report_folders(keep=dir_lvl3:);
        dir_lvl1="&forecast_report_folder.";
        product_line_group="&product_line_group.";
        seasonality="&seasonality.";
        seasonality1=tranwrd(seasonality, '-', '_');
        running_sales_week="&current_year_week.";
        material_division="&_material_division.";
        region="&region.";
        dir1=catx('_',product_line_group, seasonality1, region, material_division);
        dir_lvl2=catx('\',dir_lvl1, dir1);
        dir2=running_sales_week;
        dir_lvl3=catx('\',dir_lvl2, dir2);
        dir31='0. Capacity and Supply';
        dir32='1. PM';
        dir33='2. PM back';
        dir34='3. SM';
        dir35='4. SM back';
        dir36='5. Consolidated file';
        dir37='6. Signed-off file';
        dir_lvl31=catx('\',dir_lvl3, dir31);
        dir_lvl32=catx('\',dir_lvl3, dir32);
        dir_lvl33=catx('\',dir_lvl3, dir33);
        dir_lvl34=catx('\',dir_lvl3, dir34);
        dir_lvl35=catx('\',dir_lvl3, dir35);
        dir_lvl36=catx('\',dir_lvl3, dir36);
        dir_lvl37=catx('\',dir_lvl3, dir37);
        rc=dcreate(dir1,dir_lvl1);
        rc=dcreate(dir2,dir_lvl2);
        rc=dcreate(dir31,dir_lvl3);
        rc=dcreate(dir32,dir_lvl3);
        rc=dcreate(dir33,dir_lvl3);
        rc=dcreate(dir34,dir_lvl3);
        rc=dcreate(dir35,dir_lvl3);
        rc=dcreate(dir36,dir_lvl3);
        rc=dcreate(dir37,dir_lvl3);    
      run;
    %end;

    %if "&step1."="Y" %then %do;
      %forecast_report_step1(  Region=&region., 
                              Product_line_group=&product_line_group., 
                              material_division=%quote(&material_division.), 
                              seasonality=&seasonality., 
                              Season=&season., 
                              current_year_week=&current_year_week.,
                              Supply_vertical_file=&Supply_vertical_file.,
                              Supply_horizontal_file=&Supply_horizontal_file.,
                              Capacity_vertical_file=&Capacity_vertical_file.,
                              Capacity_horizontal_file=&Capacity_horizontal_file.,
                              previous_forecast4_sas_table=&previous_forecast4_sas_table.
                              );
    %end;

    %if "&step2."="Y" %then %do;
      %forecast_report_step2(  Region=&region., 
                              Product_line_group=&product_line_group., 
                              material_division=%quote(&material_division.), 
                              seasonality=&seasonality., 
                              Season=&season., 
                              current_year_week=&current_year_week.,
                              pm_feedback=&pm_feedback.
                              );
    %end;

    %if "&step3."="Y" %then %do;
      %forecast_report_step3(  Region=&region., 
                              Product_line_group=&product_line_group., 
                              material_division=%quote(&material_division.), 
                              seasonality=&seasonality., 
                              Season=&season., 
                              current_year_week=&current_year_week.,
                              sm_feedback_folder=%quote(&sm_feedback_folder.)
                              );
    %end;

    %if "&step4."="Y" %then %do;
      %forecast_report_step4(  Region=&region., 
                              Product_line_group=&product_line_group., 
                              material_division=%quote(&material_division.), 
                              seasonality=&seasonality., 
                              Season=&season., 
                              current_year_week=&current_year_week.,
                              signoff_file=%quote(&signoff_file.)
                              );
    %end;
  %end;

%mend forecast_reports;

%forecast_reports();
