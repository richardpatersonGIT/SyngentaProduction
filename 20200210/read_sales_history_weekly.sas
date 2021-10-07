/***********************************************************************/
/*Type: Utility*/
/*Use: Used only in forecast_report*/
/*Purpose: Create extrapolation table on Species level.*/
/*Parameters: last_season_year - year of the middle season week*/ 
/*            last_season_week - week of the middle season week*/
/*            end_season_year - year of the end season week*/
/*            end_season_week - week of the end season week*/
/*            previous_season - season for extrapolation(only first year)*/
/*            region - 'FPS' or 'BI'*/
/*            material_division - if forms like ["6A"] or like ["6B", "6C"]*/
/*            product_line_group - product line group */
/*IN: shw.shw_YYYY_WkWW_* sas datasets*/
/*OUT: work.sales_percentage*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\filter_orders.sas";

%macro read_sales_history_weekly(last_season_year=, 
                                last_season_week=, 
                                end_season_year=, 
                                end_season_week=, 
                                previous_season=, 
                                region=, 
                                material_division=, 
                                product_line_group=);

  proc datasets library=work nolist;
    delete sales_history;
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
    if (year=&last_season_year. and week=&last_season_week.) or (year=&end_season_year. and week=&end_season_week.) then output;
  run;

  data filelist3;
    set filelist2;
    order=_n_;
  run;

  proc sql noprint;
  select count(*) into : flcnt from filelist3;
  quit;
  
  %do i=1 %to &flcnt.;
    proc sql noprint;
      select dsname,year, week into :dsname trimmed, :year trimmed, :week trimmed from filelist3 where order = &i.;
    quit;

    data sales_history_raw(keep=soldto_nr sls_org sls_off shipto_cntry material variety SchedLine_Cnf_deldte cnf_qty order_type mat_div rsn_rej_cd filename);
      set shw.&dsname.;
      filename="&dsname.";
    run;

    data sales_history_tmp(drop=sls_org sls_off shipto_cntry order_type rsn_rej_cd);
      set Sales_history_raw;
      length unique_code $10. year 8. week 8.;
      unique_code=cats(sls_org, sls_off, shipto_cntry);
      year=&year.;
      week=&week.;
      if ((mat_div in ('6B', '6C') and order_type in ('ZYPD', 'ZFD1', 'ZYPL', 'ZMTO')) or (mat_div='6A' and order_type in ('YQOR', 'ZMTO'))) and missing(rsn_rej_cd) and ^missing(SchedLine_Cnf_deldte) then output;
    run;

    data sales_history_tmp1(drop=mat_div);
      length filename $100.;
      set sales_history_tmp;
      if mat_div in (&material_division.) then output;
    run;

    /*<corection of date format>*/
    proc contents data=sales_history_tmp1 out=contents noprint;
    run;

    %let datetype=1;

    proc sql noprint;
      select type into :datetype trimmed from contents where lower(name)="schedline_cnf_deldte";
    quit;

    %if "&datetype."="2" %then %do;
      data sales_history_tmp1(drop=_SchedLine_Cnf_deldte);
        set sales_history_tmp1(rename=(SchedLine_Cnf_deldte=_SchedLine_Cnf_deldte));
        SchedLine_Cnf_deldte=input(_SchedLine_Cnf_deldte, 8.);
        SchedLine_Cnf_deldte=SchedLine_Cnf_deldte-21916;  
      run;
    %end;
    /*</corection of date format>*/

    proc datasets lib=work memtype=data nolist;
       modify sales_history_tmp1;
       attrib _all_ format=;
    run;
    
    data _null_;
    sleep=sleep(5);
    run;


    proc append base=sales_history data=sales_history_tmp1;
    run;

  %end;

  data sales_history1(drop=rc);
    set sales_history;
    length region $3. territory $3. country $6.;

    if _n_=1 then do;
      declare hash cl(dataset: 'dmimport.Country_lookup');
        rc=cl.DefineKey ('unique_code');
        rc=cl.DefineData ('region', 'territory', 'country');
        rc=cl.DefineDone();
      declare hash stl(dataset: 'dmimport.Soldto_nr_lookup');
        rc=stl.DefineKey ('soldto_nr');
        rc=stl.DefineData ('region', 'territory', 'country');
        rc=stl.DefineDone();
    end;

    rc=cl.find(); /*gets territory and country from country_lookup*/
    if region='BI' then do;
      rc=stl.find();/*gets territory and country from soldto_nr_lookup (if found overwrite the country_lookup)*/
    end;
    if region="&region." then output;

  run;

  data sales_history3(drop=rc);
    set sales_history1;
    length sub_unit 8.;

    if _n_=1 then do;
      declare hash fps_assortment(dataset: 'dmimport.FPS_assortment');
        rc=fps_assortment.DefineKey ('region', 'material');
        rc=fps_assortment.DefineData ('sub_unit');
        rc=fps_assortment.DefineDone();

      declare hash bi_assortment(dataset: 'dmimport.Material_class_table');
        rc=bi_assortment.DefineKey ('region', 'material');
        rc=bi_assortment.DefineData ('sub_unit');
        rc=bi_assortment.DefineDone();
    end;

    If region = 'FPS' then rc=fps_assortment.find(); /*gets sub_unit from fps_assortment*/
    else rc=bi_assortment.find(); /*gets sub_unit from material_class_table*/

    if ^missing(sub_unit) then do;
      historical_sales=cnf_qty * sub_unit;
    end;

  run;

  data sales_history4(drop=rc);
    set sales_history3;
    length  season_week_start season_week_end 
             Order_season_start order_year order_season order_week Order_yweek order_month 8.;
    length hash_species_name $29. product_line_group $20.;
    format Order_season_start yymmdd10.;
    if _n_=1 then do;
      declare hash pmd_assortment(dataset: 'dmproc.PMD_assortment');
        rc=pmd_assortment.DefineKey ('region', 'variety');
        rc=pmd_assortment.DefineData ('season_week_start', 'season_week_end', 'hash_species_name', 'product_line_group');
        rc=pmd_assortment.DefineDone();
    end;

    rc=pmd_assortment.find(); 

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

    if order_season=&previous_season. then output;
  run;

  data sales_history5;
    set sales_history4;
    if product_line_group="&product_line_group." then output;
  run;

  %filter_orders(in_table=sales_history5, out_table=sales_history6);

  proc sql;
    create table sales_history_aggr as
    select region, year, week, hash_species_name, order_season, sum(historical_sales) as historical_sales_per_species 
      from sales_history6 
      group by region, year, week, hash_species_name, order_season;
  quit;

  data last_season_aggr(rename=(historical_sales_per_species=last_season_sales)) end_season_aggr(rename=(historical_sales_per_species=end_season_sales));
    set sales_history_aggr;
    if year=&last_season_year. and week=&last_season_week. then output last_season_aggr;
    if year=&end_season_year. and week=&end_season_week. then output end_season_aggr;
  run;

  proc sql;
  create table sales_percentage as
    select l.region, l.hash_species_name, l.last_season_sales, e.end_season_sales,  l.last_season_sales/e.end_season_sales as percentage from last_season_aggr l
    inner join end_season_aggr e 
      on l.hash_species_name=e.hash_species_name;
  quit;

%mend read_sales_history_weekly;