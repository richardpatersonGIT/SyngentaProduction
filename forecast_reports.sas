/***********************************************************************/
/*Type: Report*/
/*Use: Fill in parameters in metadata.xlsx, sheet=Forecast_report and press run*/
/*Purpose: Create forecast report with 4 seperate steps*/
/*OUT: dmfcst1.&region._&product_line_group._&ka_material_division._&first_season._&current_year_week._&_seasonality.*/
/*     dmfcst2.&region._&product_line_group._&ka_material_division._&first_season._&current_year_week._&_seasonality.*/
/*     dmfcst3.&region._&product_line_group._&ka_material_division._&first_season._&current_year_week._&_seasonality.*/
/*     dmfcst4.&region._&product_line_group._&ka_material_division._&first_season._&current_year_week._&_seasonality.*/
/*     Muptiple excel report files in in upload_report_folder (check configuration.sas)*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\cleanup_xlsx_bak_folder.sas";
%include "&sas_applications_folder.\read_sales_history_weekly.sas";
%include "&sas_applications_folder.\xls_capacity.sas";
%include "&sas_applications_folder.\xls_supply.sas";
%include "&sas_applications_folder.\concatenate_forecast_sm_feedback.sas";
%include "&sas_applications_folder.\read_metadata.sas";
%include "&sas_applications_folder.\filter_orders.sas";

%macro refresh_plc(table_in=, table_out=);

  data &table_out.(drop=rc);
    set &table_in.;
    call missing(plc, future_plc, valid_from_date, global_plc, global_future_plc, global_valid_from_date);
    if _n_=1 then do;
      declare hash refresh_plc(dataset: 'dmproc.PMD_assortment(rename=(current_plc=plc future_plc_active_date=valid_from_date  global_current_plc=global_plc global_future_plc_active_date=global_valid_from_date))');
          rc=refresh_plc.DefineKey ('region', 'variety');
          rc=refresh_plc.DefineData ('plc', 'future_plc', 'valid_from_date', 'global_plc', 'global_future_plc', 'global_valid_from_date');
          rc=refresh_plc.DefineDone();
    end;

    rc=refresh_plc.find();
  run;

%mend refresh_plc;

%macro refresh_sales(table_in=, table_out=, refresh_year_week=);
  %let _seasonality=%sysfunc(tranwrd(%quote(&seasonality.),%str(-),%str(_)));
  %let kda_material_division=%sysfunc(compress(%quote(&material_division.),,kda));
  %let ka_material_division=%sysfunc(compress(%quote(&material_division.),,ka));

  proc sql;
  create table orders_s_aggr_c as
    select region, variety, country, order_season, sum(historical_sales) as historical_sales, sum(actual_sales) as actual_sales 
      from orders_filtered
      where ^missing(order_season) and ^missing(actual_sales) and mat_div in (&material_division.) and product_line_group="&product_line_group." and region="&region." 
      group by region, variety, country, order_season;
  quit;

  proc sql;
  create table orders_s_aggr_r as
    select region, variety, region as country, order_season, sum(historical_sales) as historical_sales, sum(actual_sales) as actual_sales 
      from orders_filtered
      where ^missing(order_season) and ^missing(actual_sales) and mat_div in (&material_division.) and product_line_group="&product_line_group." and region="&region."
      group by region, variety, order_season;
  quit;

  data orders_s_aggr;
    set orders_s_aggr_c 
        orders_s_aggr_r;
  run;

  data rs(keep=region variety product_line_group current_plc seasons hash_product_line hash_species_name);
    set DMPROC.PMD_Assortment;
    length seasons $5.;
    seasons=strip(strip(put(season_week_start, 2.))||'-'||strip(put(season_week_end, 2.)));
    if region="&region.";
  run;

  data rs1;
    set rs;
      if strip(product_line_group) = "&Product_line_group." and current_plc in ('E2', 'F0', 'F1', 'F2', 'F3', 'F4', 'G0', 'G2') and strip(seasons)="&seasonality." then output;
  run;

  proc sql;
    create table regions4report as 
      select distinct 1 as order, 'Country' as type, country from (select region, country from dmimport.country_lookup union select region, country from dmimport.soldto_nr_lookup) where region="&region." and region^=country /*exception for BI/BI/BI, will also work after split into regions and countries*/
      union select distinct 3 as order, region as type, region as country from (select region, country from dmimport.country_lookup union select region, country from dmimport.soldto_nr_lookup) where region="&region.";
  quit;

  proc sql;
    create table allvarieties4report as 
    select order, type, variety, country from regions4report
    join (select distinct variety from rs1) on 1=1
    order by variety, order, country;
  quit;

  proc sql;
    create table rs2 as
    select 
      region,
      seasons, 
      country, 
      a.variety,
      hash_species_name
    from allvarieties4report a
    left join rs1 f on f.variety=a.variety;
  quit;

  %let previous_season=%eval(&season.-1);

  %let refresh_year=%substr(&refresh_year_week., 1, 4);
  %let refresh_week=%substr(&refresh_year_week., 5, 2);
  %let season_start_week=%scan(&seasonality., 1, "-");
  %let season_end_week=%scan(&seasonality., 2, "-");
  %let last_season_year=%eval(&refresh_year.-1);
  %let last_season_week=&refresh_week.;

  %if &season_end_week >=52 %then %do;
    %let end_season_year=%eval(&season.-1);
  %end; %else %do;
    %let end_season_year=&season.;
  %end;
  %let end_season_week=&season_end_week.;

  %read_sales_history_weekly(last_season_year=&last_season_year., 
                            last_season_week=&last_season_week., 
                            end_season_year=&end_season_year., 
                            end_season_week=&end_season_week., 
                            previous_season=&previous_season., 
                            region=&region., 
                            material_division=%quote(&material_division.), 
                            product_line_group=&product_line_group.);


  data rs3(keep=variety country ytd percentage extrapolation);
    length rc 8.;
    set rs2;
    length historical_sales actual_sales s1 s2 s3 order_season ytd percentage extrapolation prev_demand1 prev_demand2 prev_demand3 8.;
    
    
    if _n_=1 then do;
      declare hash aorders(dataset: 'orders_s_aggr');
        rc=aorders.DefineKey ('region', 'variety', 'order_season', 'country');
        rc=aorders.DefineData ('actual_sales');
        rc=aorders.DefineDone();
      declare hash sales_percentage(dataset: 'sales_percentage');
        rc=sales_percentage.DefineKey ('hash_species_name');
        rc=sales_percentage.DefineData ('percentage');
        rc=sales_percentage.DefineDone();
    end;

    percentage=coalesce(percentage,0);
    order_season=&season.;
    rc=aorders.find(); 
    ytd=coalesce(actual_sales,0);

    rc=sales_percentage.find();

    if percentage^=0 then do;
      extrapolation=coalesce(round(ytd/percentage, 1),0);
    end; else do;
      extrapolation=0;
    end;

  run;

  proc sql;
    create table &table_out. as
    select a.*, b.ytd, b.percentage, b.extrapolation from &table_in.(drop=ytd percentage extrapolation) a
    left join rs3 b on a.variety=b.variety and a.country=b.country;
  quit;

%mend refresh_sales;

%macro replacer_logic(table_in=, table_out=, first_season=, demand_col=, seasonality=);

  proc sql;
    create table replacer_logic1 as
    select distinct
           a.variety, a.replace_by, a.valid_from_date, a.replacement_date, a.plc, a.future_plc, 
           b.plc as replacer_plc, a.season_week_start,
           a.&demand_col.1, a.&demand_col.2, a.&demand_col.3           
    from &table_in. a 
    left join &table_in. b on a.replace_by=b.variety and a.order=b.order
    where a.order=3 
    order by a.variety, a.replace_by, a.valid_from_date, a.replacement_date, a.plc, a.future_plc, replacer_plc, a.season_week_start;
  quit;

  proc transpose data=replacer_logic1 out=replacer_logic2(rename=(_name_=_demand_season col1=demand_value));
  by variety replace_by valid_from_date replacement_date plc future_plc replacer_plc season_week_start;
  run;

  data replacer_shift;
    set replacer_logic2 (where=(^missing(replace_by)));
    format first_day_of_the_season ddmmyy10.;
    variety_shift=0;
    replacer_shift=0;
    %do season_loop=1 %to 3;
      if _demand_season="&demand_col.&season_loop." then do;
        demand_season=&first_season.+&season_loop.-1;
        first_day_of_the_season=input(put(demand_season, 4.)||'W'||put(%scan(&seasonality., 1, '-'), z2.)||'01', weekv9.);
        if ^missing(valid_from_date) and valid_from_date <= first_day_of_the_season then season_match=1; 
          else season_match=0;
        if 
           /* COMBINE FOR REPLACER LOGIC - long version
           scenario_05 (plc^='G2' and future_plc='G2' and replacer_plc^='G2' and season_match=1) or
           scenario_09 (plc='G2' and missing(future_plc) and replacer_plc^='G2' and missing(valid_from_date)) or
           scenario_10 (plc='G2' and missing(future_plc) and replacer_plc^='G2' and ^missing(valid_from_date)) or
           scenario_13 (plc='G2' and ^missing(future_plc) and future_plc^='G2' and replacer_plc^='G2' and ^missing(valid_from_date) and season_match=0) or
           scenario_14 (plc='G2' and ^missing(future_plc) and future_plc^='G2' and replacer_plc^='G2' and missing(valid_from_date)) or
           scenario_18 (plc='G2' and future_plc='G2' and replacer_plc^='G2' and season_match=1)
            */
           (future_plc='G2' and replacer_plc^='G2' and season_match=1) or
           (plc='G2' and replacer_plc^='G2' and missing(future_plc)) or
           ((plc='G2' and replacer_plc^='G2' and ^missing(future_plc) and future_plc^='G2') and 
              ((^missing(valid_from_date) and season_match=0) or
              ( missing(valid_from_date)))
           )
        then do;
          variety_shift=-demand_value;
          replacer_shift=demand_value;
        end;
      end;
    %end;
    if variety_shift^=0 then output;
  run;

  data variety_zero;
    set replacer_logic2;
    format first_day_of_the_season ddmmyy10.;
    variety_zero=0;
    %do season_loop=1 %to 3;
      if _demand_season="&demand_col.&season_loop." then do;
        demand_season=&first_season.+&season_loop.-1;
        first_day_of_the_season=input(put(demand_season, 4.)||'W'||put(%scan(&seasonality., 1, '-'), z2.)||'01', weekv9.);
        if ^missing(valid_from_date) and valid_from_date <= first_day_of_the_season then season_match=1; 
          else season_match=0;
        if 
          /*scenario_08 (plc^='G2' and future_plc='G2' and replacer_plc='G2' and season_match=1) or*/
          /*scenario_12 (plc='G2' and missing(future_plc) and replacer_plc='G2' and season_match=1) or*/
          /*scenario_13 (plc='G2' and missing(future_plc) and replacer_plc='G2' and season_match=0) or*/
          /*scenario_16 (plc='G2' and ^missing(future_plc) and future_plc^='G2' and replacer_plc='G2' and ^missing(valid_from_date) and season_match=0) or*/
          /*scenario_17 (plc='G2' and ^missing(future_plc) and future_plc^='G2' and replacer_plc='G2' and missing(valid_from_date) and season_match=0) or*/
          /*scenario_18 (plc='G2' and future_plc='G2' and replacer_plc^='G2' and ^missing(valid_from_date) and season_match=0) or*/
          /*scenario_19 (plc='G2' and future_plc='G2' and replacer_plc^='G2' and missing(valid_from_date) and season_match=0) or*/
          /*scenario_20 (plc='G2' and future_plc='G2' and replacer_plc='G2' and ^missing(valid_from_date)) or*/
          /*scenario_21 (plc='G2' and future_plc='G2' and replacer_plc='G2' and missing(valid_from_date))*/

          (plc^='G2' and future_plc='G2' and replacer_plc='G2' and season_match=1) or
          (plc='G2' and 
            ((missing(future_plc) and replacer_plc='G2') or
            (^missing(future_plc) and future_plc^='G2' and replacer_plc='G2'  and season_match=0) or
            (future_plc='G2' and replacer_plc^='G2' and season_match=0) or
            (future_plc='G2' and replacer_plc='G2'))
          )
        then do;
          variety_zero=1;
        end;
      end;
    %end;
    if variety_zero=1 then output;
  run;

  data &table_out.(drop=rc variety_shift replacer_shift variety_zero demand_season);
    set &table_in.;
    length rc variety_shift replacer_shift variety_zero demand_season 8.;
    if _n_=1 then do;
      declare hash shift_variety(dataset: 'replacer_shift');
        rc=shift_variety.DefineKey ('variety', 'demand_season');
        rc=shift_variety.DefineData ('variety_shift');
        rc=shift_variety.DefineDone();
      declare hash shift_replacer(dataset: 'replacer_shift(drop=variety rename=(replace_by=variety))', multidata:'Y');
        rc=shift_replacer.DefineKey ('variety', 'demand_season');
        rc=shift_replacer.DefineData ('replacer_shift');
        rc=shift_replacer.DefineDone();
      declare hash zero_variety(dataset: 'variety_zero');
        rc=zero_variety.DefineKey ('variety', 'demand_season');
        rc=zero_variety.DefineData ('variety_zero');
        rc=zero_variety.DefineDone();
    end;
    if order=3 then do;
      %do season_loop=1 %to 3;
        demand_season=&first_season.+&season_loop.-1;
        rc=shift_variety.find();
        if rc=0 then do;
          &demand_col.&season_loop.=&demand_col.&season_loop.+variety_shift;
        end;

        rc=shift_replacer.find();
        do while (rc=0);
          &demand_col.&season_loop.=&demand_col.&season_loop.+replacer_shift;
          rc=shift_replacer.find_next();
        end;

        rc=zero_variety.find();
        if rc=0 then do;
          &demand_col.&season_loop.=0;
        end;
      %end;
    end;
  run;

%mend replacer_logic;

%macro g2_logic(table_in=, table_out=, first_season=, demand_col=, seasonality=);

  data &table_out.(drop=first_day_of_the_season season_match demand_season);
    set &table_in.;
    format first_day_of_the_season ddmmyy10.;
    %do season_loop=1 %to 3;
      demand_season=&first_season.+&season_loop.-1;
      first_day_of_the_season=input(put(demand_season, 4.)||'W'||put(%scan(&seasonality., 1, '-'), z2.)||'01', weekv9.);
      if ^missing(valid_from_date) and valid_from_date <= first_day_of_the_season then season_match=1; 
        else season_match=0;
      if 
         /* COMBINE FOR G2 LOGIC - long version
          scenario_01/02 (plc='G2' and future_plc='G2') or
          scenario_03/04 (plc='G2' and future_plc^='G2' and season_match=0) or
          scenario_05/06 (plc='G2' and missing(future_plc)) or
          scenario_07    (plc^='G2' and future_plc='G2' and season_match=1)
          */
        (plc='G2' and (
          (future_plc='G2') or
          (future_plc^='G2' and season_match=0) or
          (missing(future_plc))
        )) or
        (plc^='G2' and future_plc='G2' and season_match=1)
      then do;
        &demand_col.&season_loop.=0;
      end;
    %end;
  run;

%mend g2_logic;

%macro e2_logic(table_in=, table_out=, first_season=, demand_col=, seasonality=);

  %if %sysfunc(find(%sysfunc(lowcase(&product_line_group.)),cu)) > 0 %then %do;

    data &table_out.(drop=first_day_of_the_season season_match demand_season);
      set &table_in.;
      format first_day_of_the_season ddmmyy10.;
      %do season_loop=1 %to 3;
        demand_season=&first_season.+&season_loop.-1;
        first_day_of_the_season=input(put(demand_season, 4.)||'W'||put(%scan(&seasonality., 1, '-'), z2.)||'01', weekv9.);
        if ^missing(valid_from_date) and valid_from_date <= first_day_of_the_season then season_match=1; 
          else season_match=0;
        if 
           /* COMBINE FOR E2 LOGIC - long version
            scenario_01/02 (plc='E2' and future_plc='E2') or
            scenario_03/04 (plc='E2' and future_plc^='E2' and season_match=0) or
            scenario_05/06 (plc='E2' and missing(future_plc)) or
            scenario_07    (plc^='E2' and future_plc='E2' and season_match=1)
            */
          (plc='E2' and (
            (future_plc='E2') or
            (future_plc^='E2' and season_match=0) or
            (missing(future_plc))
          )) or
          (plc^='E2' and future_plc='E2' and season_match=1)
        then do;
          if region='SFE' then do;
             if &demand_col.&season_loop.^=0 and ^missing(&demand_col.&season_loop.) then do;
              &demand_col.&season_loop.=max(&demand_col.&season_loop., 40000);
             end;
          end; else do;
             if &demand_col.&season_loop.^=0 and ^missing(&demand_col.&season_loop.) then do;
              &demand_col.&season_loop.=max(&demand_col.&season_loop., 10000);
             end;
          end;
        end;
      %end;
    run;

  %end; %else %do;

    data &table_out.;
      set &table_in.;
    run;

  %end;

%mend e2_logic;

%macro forecast_report_step1(Round=, refresh_plc_from_PMD=, Region=, Product_line_group=, material_division=, seasonality=, Season=, First_season=, base_season_yyyy=, base_season_x=, current_year_week=, Supply_vertical_file=, Supply_horizontal_file=, Capacity_vertical_file=, Capacity_horizontal_file=, previous_forecast_file=);

  %let _seasonality=%sysfunc(compress(%substr(%quote(&seasonality.)., 1, 2),,kd));
  %let kda_material_division=%sysfunc(compress(%quote(&material_division.),,kda));
  %let excel_report_fname=&region._&product_line_group._&kda_material_division._&first_season._&current_year_week._&_seasonality.;
  %let ka_material_division=%sysfunc(compress(%quote(&material_division.),,ka));
  %let sas_report_fname=&region.&product_line_group.&ka_material_division.&first_season.%substr(&round., 1, 1)&current_year_week.&_seasonality.;

  proc sql;
  create table orders_s_aggr_c as
    select region, variety, country, order_season, sum(historical_sales) as historical_sales, sum(actual_sales) as actual_sales 
      from orders_filtered
      where ^missing(order_season) and ^missing(actual_sales) and mat_div in (&material_division.) and product_line_group="&product_line_group." and region="&region." 
      group by region, variety, country, order_season;
  quit;

  proc sql;
  create table orders_s_aggr_r as
    select region, variety, region as country, order_season, sum(historical_sales) as historical_sales, sum(actual_sales) as actual_sales 
      from orders_filtered
      where ^missing(order_season) and ^missing(actual_sales) and mat_div in (&material_division.) and product_line_group="&product_line_group." and region="&region."
      group by region, variety, order_season;
  quit;

  data orders_s_aggr;
    set orders_s_aggr_c 
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
      if strip(product_line_group) = "&Product_line_group." and current_plc in ('E2', 'F0', 'F1', 'F2', 'F3', 'F4', 'G0', 'G2') and strip(seasons)="&seasonality." then output;
  run;

  proc sql;
    create table regions4report as 
      select distinct 1 as order, 'Country' as type, country from (select region, country from dmimport.country_lookup union select region, country from dmimport.soldto_nr_lookup) where region="&region." and region^=country /*exception for BI/BI/BI, will also work after split into regions and countries*/
      union select distinct 3 as order, region as type, region as country from (select region, country from dmimport.country_lookup union select region, country from dmimport.soldto_nr_lookup) where region="&region.";
  quit;

  proc sql;
    create table allvarieties4report as 
    select order, type, variety, country from regions4report
    join (select distinct variety from fr1) on 1=1
    order by variety, order, country;
  quit;

  proc sql;
    create table fr2 as
    select 
      order,
      region,
      type, 
      seasons, 
      country, 
      "" as crop_categories length=18,
      genetics,
      species_code,
      product_line,
      species,
      series,
      a.variety,
      variety_name,
      current_plc as plc,
      future_plc,
      future_plc_active_date as valid_from_date,
      replace_by,
      replacement_date,
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
        rc=horders.DefineKey ('region', 'variety', 'order_season', 'country');
        rc=horders.DefineData ('historical_sales');
        rc=horders.DefineDone();
      declare hash aorders(dataset: 'orders_s_aggr');
        rc=aorders.DefineKey ('region', 'variety', 'order_season', 'country');
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

    call missing(historical_sales);
    order_season=&season4.;
    rc=horders.find();
    s4=coalesce(historical_sales,0);


	call missing(historical_sales);
    order_season=&season5.;
    rc=horders.find();
    s5=coalesce(historical_sales,0);

    rc=sales_percentage.find();

    if percentage^=0 then do;
      extrapolation=coalesce(round(ytd/percentage, 1),0);
    end; else do;
      extrapolation=0;
    end;

  run;

  proc sql noprint;
    select 
      case when sum(a.s3)^=0 then (sum(a.s1)/sum(a.s3))**(1/2)-1
        else 1
      end as CAGR into :cagr
      from fr3 a 
      left join dmimport.tactical_plan b on a.region=b.region and a.product_line=b.product_line and a.species_code=b.species_code and b.hash_mat_div="&ka_material_division."
      where missing(b.crop_categories) and a.order=3;
  quit;

  data fr3 (drop=rc tp_growth tactical_plan_base tp_spread);
    set fr3;
    length rc 8.;
    length hash_mat_div $4. tp_spread $9. tp_growth 8. perc_growth1 perc_growth2 perc_growth3 tactical_plan_base tactical_plan1 tactical_plan2 tactical_plan3 8.;

    hash_mat_div="&ka_material_division.";

    if _n_=1 then do;
      declare hash tactical_plan(dataset: 'dmimport.tactical_plan');
          rc=tactical_plan.DefineKey ('region', 'product_line', 'species_code', 'hash_mat_div', 'country', 'tp_spread');
          rc=tactical_plan.DefineData ('tp_growth');
          rc=tactical_plan.DefineDone();
      declare hash tactical_plan_region(dataset: 'dmimport.tactical_plan');
          rc=tactical_plan_region.DefineKey ('region', 'product_line', 'species_code', 'hash_mat_div');
          rc=tactical_plan_region.DefineData ('crop_categories');
          rc=tactical_plan_region.DefineDone();
    end;

      %if "&base_season_x."="H" %then %do;
        tactical_plan_base=s1;
      %end; %else %if "&base_season_x."="Y" %then %do;
        tactical_plan_base=ytd;
      %end; %else %if "&base_season_x."="E" %then %do;
        tactical_plan_base=extrapolation;
      %end;

    rc=tactical_plan_region.find();
    if rc^=0 then do;
      crop_categories="Service";
      if order=1 then do;
        perc_growth1=&cagr.;
        perc_growth2=&cagr.;
        perc_growth3=&cagr.;
        tactical_plan1=tactical_plan_base*(1+coalesce(perc_growth1, 0))**2;
        tactical_plan2=tactical_plan1    *(1+coalesce(perc_growth2, 0));
        tactical_plan3=tactical_plan2    *(1+coalesce(perc_growth3, 0));
      end;
    end; else do;

      if order=1 then do;
        tp_spread="&base_season_yyyy.-&first_season.";
        rc=tactical_plan.find();
        if rc=0 then perc_growth1=tp_growth;

        tp_spread="%eval(&first_season.)-%eval(&first_season.+1)";
        rc=tactical_plan.find();
        if rc=0 then perc_growth2=tp_growth;

        tp_spread="%eval(&first_season.+1)-%eval(&first_season.+2)";
        rc=tactical_plan.find();
        if rc=0 then perc_growth3=tp_growth;

  /*      if missing(replacement_date) then do;*/
  /*        if ^missing(replace_by) and ^missing(valid_from_date) then do;*/
  /*          replacement_date=valid_from_date;*/
  /*        end;*/
  /*      end;*/

        tactical_plan1=tactical_plan_base*(1+coalesce(perc_growth1, 0));
        tactical_plan2=tactical_plan1    *(1+coalesce(perc_growth2, 0));
        tactical_plan3=tactical_plan2    *(1+coalesce(perc_growth3, 0));
      end;
    end;

    tactical_plan1=max(tactical_plan1, 0);
    tactical_plan2=max(tactical_plan2, 0);
    tactical_plan3=max(tactical_plan3, 0);

  run;

  proc sort data=fr3 out=fr3; 
    by variety order region country ;
  run;

  data fr3;
    set fr3;
    by variety order region country;
    length region_tactical_plan1 region_tactical_plan2 region_tactical_plan3 8.;
    retain region_tactical_plan1 region_tactical_plan2 region_tactical_plan3;
    if first.variety then do;
      region_tactical_plan1=0;
      region_tactical_plan2=0;
      region_tactical_plan3=0;
    end;
    if order=1 then do;
      region_tactical_plan1=region_tactical_plan1+tactical_plan1;
      region_tactical_plan2=region_tactical_plan2+tactical_plan2;
      region_tactical_plan3=region_tactical_plan3+tactical_plan3;
    end;
    if order=3 then do;
      tactical_plan1=region_tactical_plan1;
      tactical_plan2=region_tactical_plan2;
      tactical_plan3=region_tactical_plan3;
    end;
  run;

  data fr3(drop=rc);
    set fr3;
    length rc 8.;
    length price 8.;
    if _n_=1 then do;
      declare hash price_list(dataset: 'dmimport.price_list');
          rc=price_list.DefineKey ('region', 'product_line', 'species_code', 'hash_mat_div');
          rc=price_list.DefineData ('price');
          rc=price_list.DefineDone();
    end;
    rc=price_list.find();
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
    order by variety, order, country;
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
    order by variety, order, country;
  quit;

  data fr6;
    set fr5;
    current_season_split=coalesce(current_season_split,0);
    last_season_split=coalesce(last_season_split,0);
  run;

  %if "&previous_forecast_file."^="" %then %do;
    %read_forecast_file(forecast_file=&previous_forecast_file., forecast_sheet=Variety level fcst);
  %end;

  data fr7(drop=rc ff_season ff_sm ff_assumption);
  length rc ff_season ff_previous_demand prev_demand0 prev_demand1 prev_demand2 prev_demand3 8.;
  length ff_assumption assumption0 assumption1 assumption2 assumption3 $1000.;
    set fr6;
    %if "&previous_forecast_file."^="" %then %do;
      if _n_=1 then do;
        declare hash hash_ff_sm (dataset: "ff_sm(rename=(season=ff_season sm=ff_sm))");
          rc=hash_ff_sm.DefineKey ('variety', 'country', 'ff_season');
          rc=hash_ff_sm.DefineData ('ff_sm');
          rc=hash_ff_sm.DefineDone();
        declare hash ff_assmt (dataset: "ff_assumptions(rename=(season=ff_season assumption=ff_assumption))");
          rc=ff_assmt.DefineKey ('variety', 'country', 'ff_season');
          rc=ff_assmt.DefineData ('ff_assumption');
          rc=ff_assmt.DefineDone();
      end;

      ff_season=&season.;
      rc=hash_ff_sm.find();
      prev_demand0=ff_sm;
      rc=ff_assmt.find();
      assumption0=ff_assumption;

      ff_season=&first_season.;
      rc=hash_ff_sm.find();
      prev_demand1=ff_sm;
      rc=ff_assmt.find();
      assumption1=ff_assumption;

      ff_season=%eval(&first_season.+1);
      rc=hash_ff_sm.find();
      prev_demand2=ff_sm;
      rc=ff_assmt.find();
      assumption2=ff_assumption;

      ff_season=%eval(&first_season.+2);
      rc=hash_ff_sm.find();
      prev_demand3=ff_sm;
      rc=ff_assmt.find();
      assumption3=ff_assumption;

    %end;

  run;

  %if %sysfunc(find(&product_line_group.,Cut)) ge 0 %then %let is_cutting=Y;
    %else %let is_cutting=N;

  data fr8;
    set fr7;
    length pm_demand1 pm_demand2 pm_demand3 8.;
    %if "&round."="REVIEW" %then %do;
      pm_demand1=prev_demand1;
      pm_demand2=prev_demand2;
      pm_demand3=prev_demand3;
    %end; 
    %if "&round."="NEW" %then %do;
      pm_demand1=tactical_plan1;
      pm_demand2=tactical_plan2;
      pm_demand3=tactical_plan3;
    %end;
    if country^="&region." then do;
      pm_demand1=0; 
      pm_demand2=0; 
      pm_demand3=0;
    end; 
  run;

  %replacer_logic(table_in=fr8, table_out=fr88, first_season=&first_season., demand_col=pm_demand, seasonality=&seasonality.);
  %g2_logic(table_in=fr88, table_out=fr888, first_season=&first_season., demand_col=pm_demand, seasonality=&seasonality.);
  %e2_logic(table_in=fr888, table_out=fr9, first_season=&first_season., demand_col=pm_demand, seasonality=&seasonality.);


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
    by product_line species series variety order country;
  run;

  %let FR_COLUMNS_FOR_TEMPLATE=order seasons Country Product_Line 
       crop_categories genetics species_code
       Species Series variety Variety_name 
       plc Future_PLC valid_from_date 
       replace_by replacement_date
       global_plc Global_Future_PLC global_valid_from_date 
       s5 s4 s3 s2 s1 
       ytd percentage extrapolation 
       country_split supply capacity remark 
       prev_demand0 assumption0
       perc_growth1 tactical_plan1 pm_demand1 sm_demand1 prev_demand1 assumption1 
       perc_growth2 tactical_plan2 pm_demand2 sm_demand2 prev_demand2 assumption2 
       perc_growth3 tactical_plan3 pm_demand3 sm_demand3 prev_demand3 assumption3
	   perc_growth4 tactical_plan4 pm_demand4 sm_demand4 prev_demand4 assumption4
	   perc_growth5 tactical_plan5 pm_demand5 sm_demand5 prev_demand5 assumption5
       price adj_percentage;

  data FOR_TEMPLATE(keep=&FR_COLUMNS_FOR_TEMPLATE.);
    retain &FR_COLUMNS_FOR_TEMPLATE.;
    length sm_demand1 sm_demand2 sm_demand3 sm_demand4 sm_demand5 8.;
    length remark $1.;
    set fr_end;
    call missing(remark);
    percentage=percentage; 
    country_split=country_split; 
    pm_demand1=round(pm_demand1,1);
    pm_demand2=round(pm_demand2,1);
    pm_demand3=round(pm_demand3,1);
	pm_demand4=0;
	pm_demand5=0;
    prev_demand1=round(prev_demand1,1);
    prev_demand2=round(prev_demand2,1);
    prev_demand3=round(prev_demand3,1);
	prev_demand4=0;
	prev_demand5=0;
    sm_demand1=round(sm_demand1,1);
    sm_demand2=round(sm_demand2,1);
    sm_demand3=round(sm_demand3,1);
	sm_demand4=0;
	sm_demand5=0;

	s5=round(s5,1);
	s4=round(s4,1);
    s3=round(s3,1);
    s2=round(s2,1);
    s1=round(s1,1);
    ytd=round(ytd,1);
    extrapolation=round(extrapolation,1);
  run;

  proc sort data=FOR_TEMPLATE;
    by product_line species series variety order country;
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

%macro forecast_report_step2(Region=, Product_line_group=, material_division=, seasonality=, Season=, current_year_week=, pm_feedback=, refresh_plc_from_PMD=, REFRESH_SALES_WEEK=);

  %let _seasonality=%sysfunc(compress(%substr(%quote(&seasonality.)., 1, 2),,kd));
  %let kda_material_division=%sysfunc(compress(%quote(&material_division.),,kda));
  %let excel_report_fname=&region._&product_line_group._&kda_material_division._&season._&current_year_week._&_seasonality.;
  %let ka_material_division=%sysfunc(compress(%quote(&material_division.),,ka));
  %let sas_report_fname=&region.&product_line_group.&ka_material_division.&first_season.%substr(&round., 1, 1)&current_year_week.&_seasonality.;

  %read_forecast_file(forecast_file=&pm_feedback., forecast_sheet=Variety level fcst);

  data pm_feedback1 (keep=variety country PLC future_plc valid_from_date global_plc global_future_plc global_valid_from_date sum_of_3_seasons pmf_demand1 pmf_demand2 pmf_demand3 pmf_assm1 pmf_assm2 pmf_assm3);
    set forecast_file(rename=(
                      dmpm_demand_&nextseason1.=pmf_demand1
                      dmpm_demand_&nextseason2.=pmf_demand2
                      dmpm_demand_&nextseason3.=pmf_demand3
                      assumptions_&nextseason1.=pmf_assm1
                      assumptions_&nextseason2.=pmf_assm2
                      assumptions_&nextseason3.=pmf_assm3
                     ));
    if missing(pmf_demand1) then pmf_demand1=0;
    if missing(pmf_demand2) then pmf_demand2=0;
    if missing(pmf_demand3) then pmf_demand3=0;
    sum_of_3_seasons=pmf_demand1+pmf_demand2+pmf_demand3;
    if country="&region." then output;
  run;

  %if "&refresh_plc_from_PMD."="Y" %then %do;
    %refresh_plc(table_in=pm_feedback1, table_out=pm_feedback2);
  %end; %else %do;
    data pm_feedback2(drop=rc);
      set pm_feedback1;
      if _N_=1 then do;
        declare hash refresh_plc(dataset: 'ff_plc');
          rc=refresh_plc.DefineKey ('variety');
          rc=refresh_plc.DefineData ('plc', 'future_plc', 'valid_from_date', 'global_plc', 'global_future_plc', 'global_valid_from_date');
          rc=refresh_plc.DefineDone();
      end;
      rc=refresh_plc.find();
    run;
  %end;

  %g2_logic(table_in=pm_feedback2, table_out=pm_feedback3, first_season=&first_season., demand_col=pmf_demand, seasonality=&seasonality.);
  %e2_logic(table_in=pm_feedback3, table_out=pm_feedback4, first_season=&first_season., demand_col=pmf_demand, seasonality=&seasonality.);

  proc sql;
    create table pmf1 as
    select a.*, 
          b.PLC,
          b.future_plc, 
          b.valid_from_date, 
          b.global_plc, 
          b.global_future_plc, 
          b.global_valid_from_date,
          b.pmf_demand1,
          b.pmf_demand2,
          b.pmf_demand3,
          b.pmf_assm1,
          b.pmf_assm2,
          b.pmf_assm3,
          round((b.pmf_demand1*a.country_split),1) as pmf_split_demand1, 
          round((b.pmf_demand2*a.country_split),1) as pmf_split_demand2, 
          round((b.pmf_demand3*a.country_split),1) as pmf_split_demand3 
    from dmfcst1.&sas_report_fname.(drop=PLC future_plc valid_from_date global_plc global_future_plc global_valid_from_date) a 
    left join pm_feedback4 b on a.variety=b.variety;
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

  data pmf4;
    set pmf3;
    length pmf_sm_demand1 pmf_sm_demand2 pmf_sm_demand3 8.;
    pmf_sm_demand1=pmf_split_demand1;
    pmf_sm_demand2=pmf_split_demand2;
    pmf_sm_demand3=pmf_split_demand3;
  run;

  %if "&refresh_sales_week."^="" %then %do;
    %refresh_sales(table_in=pmf4, table_out=pmf5, refresh_year_week=&refresh_sales_week.);
  %end; %else %do;
    data pmf5;
      set pmf4;
    run;
  %end;

  data pmf_end;
    set pmf5;
  run;

  proc sort data=pmf_end;
    by product_line species series variety order country;
  run; 

  data dmfcst2.&sas_report_fname.;
    set pmf_end;
  run;

  %let PMF_COLUMNS_FOR_TEMPLATE=order seasons Country Product_Line 
        crop_categories genetics species_code
        Species Series variety Variety_name 
        plc Future_PLC valid_from_date 
        replace_by replacement_date
        global_plc Global_Future_PLC global_valid_from_date 
        s3 s2 s1 
        ytd percentage extrapolation 
        country_split pmf_supply pmf_capacity remark 
        prev_demand0 assumption0
        perc_growth1 tactical_plan1 pmf_split_demand1 pmf_sm_demand1 prev_demand1 pmf_assm1
        perc_growth2 tactical_plan2 pmf_split_demand2 pmf_sm_demand2 prev_demand2 pmf_assm2 
        perc_growth3 tactical_plan3 pmf_split_demand3 pmf_sm_demand3 prev_demand3 pmf_assm3 
        price;

  data FOR_TEMPLATE_COUNTRIES(keep=&PMF_COLUMNS_FOR_TEMPLATE.);
    retain &PMF_COLUMNS_FOR_TEMPLATE.;
    set pmf_end;
  call missing(remark);
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
    by product_line species series variety order country;
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

%macro forecast_report_step3(Region=, Product_line_group=, material_division=, seasonality=, Season=, current_year_week=, sm_feedback_folder=, refresh_plc_from_PMD=, REFRESH_SALES_WEEK=);

  %let _seasonality=%sysfunc(compress(%substr(%quote(&seasonality.)., 1, 2),,kd));
  %let kda_material_division=%sysfunc(compress(%quote(&material_division.),,kda));
  %let excel_report_fname=&region._&product_line_group._&kda_material_division._&season._&current_year_week._&_seasonality.;
  %let ka_material_division=%sysfunc(compress(%quote(&material_division.),,ka));
  %let sas_report_fname=&region.&product_line_group.&ka_material_division.&first_season.%substr(&round., 1, 1)&current_year_week.&_seasonality.;

  %concatenate_forecast_sm_feedback(sm_feedback_folder=%quote(&sm_feedback_folder.));

  data smf1;
    set dmfcst2.&sas_report_fname.;
  run;

  %if "&refresh_plc_from_PMD."="Y" %then %do;
    %refresh_plc(table_in=smf1, table_out=smf1);
  %end;

  proc sql;
    create table forecast_sm_feedback_demand1 as
    select country, variety, smf_demand1, smf_demand2, smf_demand3 
  from forecast_sm_feedback_demand 
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
    smf_total_demand_raw1=smf_total_demand1;
    smf_total_demand_raw2=smf_total_demand2;
    smf_total_demand_raw3=smf_total_demand3;
  run;
  
  %g2_logic(table_in=smf3, table_out=smf4, first_season=&first_season., demand_col=smf_total_demand, seasonality=&seasonality.);
  %e2_logic(table_in=smf4, table_out=smf5, first_season=&first_season., demand_col=smf_total_demand, seasonality=&seasonality.);

  data smf6;
    set smf5;
    if smf_total_demand_raw1^=0 then smf_demand1=round((smf_demand1_raw/smf_total_demand_raw1)*smf_total_demand1, 1);
      else smf_demand1=0;
    if smf_total_demand_raw2^=0 then smf_demand2=round((smf_demand2_raw/smf_total_demand_raw2)*smf_total_demand2, 1);
      else smf_demand2=0;
    if smf_total_demand_raw3^=0 then smf_demand3=round((smf_demand3_raw/smf_total_demand_raw3)*smf_total_demand3, 1);
      else smf_demand3=0;
  run;

  %if "&refresh_sales_week."^="" %then %do;
    %refresh_sales(table_in=smf6, table_out=smf7, refresh_year_week=&refresh_sales_week.);
  %end; %else %do;
    data smf7;
      set smf6;
    run;
  %end;
  
  data smf_end;
    set smf7;
  run;

  proc sort data=smf_end;
    by product_line species series variety order country;
  run; 

  data dmfcst3.&sas_report_fname.;
    set smf_end;
  run;

  %let SMF_COLUMNS_FOR_TEMPLATE=order seasons Country Product_Line 
     crop_categories genetics species_code
     Species Series variety Variety_name 
     plc Future_PLC valid_from_date 
     replace_by replacement_date
     global_plc Global_Future_PLC global_valid_from_date  
     s3 s2 s1  
     ytd percentage extrapolation  
     country_split pmf_supply pmf_capacity remark  
     prev_demand0 assumption0
     perc_growth1 tactical_plan1 pmf_split_demand1 smf_demand1 prev_demand1 smf_assm1
     perc_growth2 tactical_plan2 pmf_split_demand2 smf_demand2 prev_demand2 smf_assm2 
     perc_growth3 tactical_plan3 pmf_split_demand3 smf_demand3 prev_demand3 smf_assm3 
     price;

  data SMF_FOR_TEMPLATE(keep=&SMF_COLUMNS_FOR_TEMPLATE.);
    retain &SMF_COLUMNS_FOR_TEMPLATE.;
    length remark $1.;
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
    by product_line species series variety order country;
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

%macro forecast_report_step4(Region=, Product_line_group=, material_division=, seasonality=, Season=, current_year_week=, signoff_file=, refresh_plc_from_PMD=, REFRESH_SALES_WEEK=);

  %let _seasonality=%sysfunc(compress(%substr(%quote(&seasonality.)., 1, 2),,kd));
  %let kda_material_division=%sysfunc(compress(%quote(&material_division.),,kda));
  %let excel_report_fname=&region._&product_line_group._&kda_material_division._&season._&current_year_week._&_seasonality.;
  %let ka_material_division=%sysfunc(compress(%quote(&material_division.),,ka));
  %let sas_report_fname=&region.&product_line_group.&ka_material_division.&first_season.%substr(&round., 1, 1)&current_year_week.&_seasonality.;

  %read_forecast_file(forecast_file=&signoff_file., forecast_sheet=Variety level fcst);

  data signoff_file0 (keep=variety country PLC so_demand1 so_demand2 so_demand3 so_assm1 so_assm2 so_assm3);
    set forecast_file(rename=(sm_demand_&nextseason1.=so_demand1 
                              sm_demand_&nextseason2.=so_demand2
                              sm_demand_&nextseason3.=so_demand3
                              assumptions_&nextseason1.=so_assm1
                              assumptions_&nextseason2.=so_assm2
                              assumptions_&nextseason3.=so_assm3));
    if missing(so_demand1) then so_demand1=0;
    if missing(so_demand2) then so_demand2=0;
    if missing(so_demand3) then so_demand3=0;
  run;

  data so1;
    set dmfcst3.&sas_report_fname.;
  run;

  proc sql;
    create table so2 as
    select a.*, b.so_demand1, b.so_demand2, b.so_demand3, b.so_assm1, b.so_assm2, b.so_assm3 
    from so1 a
    left join signoff_file0 b on a.variety=b.variety and a.country=b.country;
  quit;

  %if "&refresh_plc_from_PMD."="Y" %then %do;
    %refresh_plc(table_in=so2, table_out=so3);
  %end; %else %do;
    data so3(drop=rc);
      set so2;
      if _N_=1 then do;
        declare hash refresh_plc(dataset: 'ff_plc');
          rc=refresh_plc.DefineKey ('variety');
          rc=refresh_plc.DefineData ('plc', 'future_plc', 'valid_from_date', 'global_plc', 'global_future_plc', 'global_valid_from_date');
          rc=refresh_plc.DefineDone();
      end;
      rc=refresh_plc.find();
    run;
  %end;

  %if "&refresh_sales_week."^="" %then %do;
    %refresh_sales(table_in=so3, table_out=so4, refresh_year_week=&refresh_sales_week.);
  %end; %else %do;
    data so4;
      set so3;
    run;
  %end;

  data so_end;
    set so4;
  run;

  proc sort data=so_end;
    by product_line species series variety order country;
  run; 

  data dmfcst4.&sas_report_fname.;
    set so_end;
  run;

  %let SO_COLUMNS_FOR_TEMPLATE=order seasons Country Product_Line 
     crop_categories genetics species_code
     Species Series variety Variety_name 
     plc Future_PLC valid_from_date 
     replace_by replacement_date
     global_plc Global_Future_PLC global_valid_from_date  
     s3 s2 s1  
     ytd percentage extrapolation  
     country_split pmf_supply pmf_capacity remark  
     prev_demand0 assumption0
     perc_growth1 tactical_plan1 so_split_demand1 so_demand1 prev_demand1 so_assm1
     perc_growth2 tactical_plan2 so_split_demand2 so_demand2 prev_demand2 so_assm2 
     perc_growth3 tactical_plan3 so_split_demand3 so_demand3 prev_demand3 so_assm3 
     price;

  data SO_FOR_TEMPLATE(keep=&SO_COLUMNS_FOR_TEMPLATE.);
    retain &SO_COLUMNS_FOR_TEMPLATE.;
    length remark $1.;
    set so_end;
    percentage=percentage; 
    country_split=country_split; 
    so_split_demand1=round(pmf_split_demand1,1);
    so_split_demand2=round(pmf_split_demand2,1);
    so_split_demand3=round(pmf_split_demand3,1);
    prev_demand1=round(prev_demand1,1);
    prev_demand2=round(prev_demand2,1);
    prev_demand3=round(prev_demand3,1);
    so_demand1=round(so_demand1,1);
    so_demand2=round(so_demand2,1);
    so_demand3=round(so_demand3,1);
    s3=round(s3,1);
    s2=round(s2,1);
    s1=round(s1,1);    
    ytd=round(ytd,1);
    extrapolation=round(extrapolation,1);
  run;

  proc sql noprint;
    select dir_lvl37 into :signoff_folder trimmed from forecast_report_folders;
  quit;

    proc export 
    data=so_end 
    dbms=xlsx 
    outfile="&signoff_folder.\&excel_report_fname._STEP4.xlsx" 
    replace;
    SHEET="Forecast_step4"; 
  run;

  proc export 
    data=SO_FOR_TEMPLATE 
    dbms=xlsx 
    outfile="&signoff_folder.\&excel_report_fname._FOR_TEMPLATE_STEP4.xlsx" 
    replace;
    SHEET="Forecast_step4"; 
  run;


%mend forecast_report_step4;

%macro forecast_reports();
  
  %read_metadata(sheet=Forecast_reports);

  /*Filter the forecast report metadata file - process only items to do*/
  data dmimport.Forecast_reports_md;
    set dmimport.Forecast_reports_md;
    if strip(step1)='Y' or
       strip(step2)='Y' or
       strip(step3)='Y' or
       strip(step4)='Y' then do;
      output;
    end;
  run;

  proc sql noprint;
    select count(*) into :report_cnt from dmimport.Forecast_reports_md;
  quit;
  
  %do ii=1 %to &report_cnt.;

    data _null_;
      set dmimport.Forecast_reports_md;
      if _n_=&ii then do;
        call symput('round', strip(round));
        call symput('refresh_plc_from_PMD', strip(refresh_plc_from_PMD));
        call symput('refresh_sales_week', strip(refresh_sales_week));
        call symput('region', strip(region));
        call symput('product_line_group', strip(product_line_group));
        call symput('outlicensing', strip(outlicensing));
        call symput('material_division', '"'||strip(tranwrd(material_division, ',', '", "'))||'"');
        call symput('season', strip(season));
        call symput('first_season', strip(first_season));
        call symput('base_season_yyyy', strip(compress(base_season_for_growth, '', 'kd')));
        call symput('base_season_x', strip(compress(base_season_for_growth, '', 'ka')));
        call symput('current_year_week', strip(current_year_week));
        call symput('seasonality', strip(seasonality));
        call symput('supply_file', strip(supply_file));
        call symput('Supply_vertical_file', strip(Supply_vertical_file));
        call symput('Supply_horizontal_file', strip(Supply_horizontal_file));
        call symput('Capacity_vertical_file', strip(Capacity_vertical_file));
        call symput('Capacity_horizontal_file', strip(Capacity_horizontal_file));
        call symput('mat_div', strip(mat_div));
        call symput('previous_forecast_file', strip(previous_forecast_file));
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
	  %let season4=%eval(&season.-4);
	  %let season5=%eval(&season.-5);
      %let nextseason1=%eval(&first_season.+0);
      %let nextseason2=%eval(&first_season.+1);
      %let nextseason3=%eval(&first_season.+2);
	  %let nextseason4=%eval(&first_season.+3);
	  %let nextseason5=%eval(&first_season.+4);

      %let _material_division=%sysfunc(compress(%quote(&material_division.),,kda));

      %put Report_variables &region. &product_line_group. &material_division. &product_line_group. &outlicensing. &season. &current_year_week. &seasonality. &mat_div. &previous_forecast_file. &step1.;
    
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

    %filter_orders(in_table=dmproc.orders_all, out_table=orders_filtered);

    %if "&step1."="Y" %then %do;
      %forecast_report_step1(  
                              Round=&round.,
                              refresh_plc_from_PMD=&refresh_plc_from_PMD.,
                              Region=&region., 
                              Product_line_group=&product_line_group., 
                              material_division=%quote(&material_division.), 
                              seasonality=&seasonality., 
                              Season=&season., 
                              First_season=&first_season., 
                              base_season_yyyy=&base_season_yyyy.,
                              base_season_X=&base_season_x.,
                              current_year_week=&current_year_week.,
                              Supply_vertical_file=&Supply_vertical_file.,
                              Supply_horizontal_file=&Supply_horizontal_file.,
                              Capacity_vertical_file=&Capacity_vertical_file.,
                              Capacity_horizontal_file=&Capacity_horizontal_file.,
                              previous_forecast_file=&previous_forecast_file.
                              );
    %end;

    %if "&step2."="Y" %then %do;
      %forecast_report_step2(  Region=&region., 
                              Product_line_group=&product_line_group., 
                              material_division=%quote(&material_division.), 
                              seasonality=&seasonality., 
                              Season=&season., 
                              current_year_week=&current_year_week.,
                              pm_feedback=&pm_feedback.,
                              refresh_plc_from_PMD=&refresh_plc_from_PMD.,
                              refresh_sales_week=&refresh_sales_week.
                              );
    %end;

    %if "&step3."="Y" %then %do;
      %forecast_report_step3(  Region=&region., 
                              Product_line_group=&product_line_group., 
                              material_division=%quote(&material_division.), 
                              seasonality=&seasonality., 
                              Season=&season., 
                              current_year_week=&current_year_week.,
                              sm_feedback_folder=%quote(&sm_feedback_folder.),
                              refresh_plc_from_PMD=&refresh_plc_from_PMD.,
                              refresh_sales_week=&refresh_sales_week.
                              );
    %end;

    %if "&step4."="Y" %then %do;
      %forecast_report_step4(  Region=&region., 
                              Product_line_group=&product_line_group., 
                              material_division=%quote(&material_division.), 
                              seasonality=&seasonality., 
                              Season=&season., 
                              current_year_week=&current_year_week.,
                              signoff_file=%quote(&signoff_file.),
                              refresh_plc_from_PMD=&refresh_plc_from_PMD.,
                              refresh_sales_week=&refresh_sales_week.
                              );
    %end;
  %end;

%mend forecast_reports;

%forecast_reports();