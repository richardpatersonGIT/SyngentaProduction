%macro extrapolation_extraction(extrapolation_config_ds=);

  proc sql noprint;
    select hist_season into :extrapolation_season trimmed from &extrapolation_config_ds. where ^missing(hist_season);
    select region into :region trimmed from &extrapolation_config_ds. where ^missing(region);
    select '"'||mat_div||'"' into :mat_div_quote separated by ', ' from &extrapolation_config_ds. where ^missing(mat_div);
    select mat_div into :mat_div separated by ',' from &extrapolation_config_ds. where ^missing(mat_div);
    select product_line_group into :product_line_group trimmed from &extrapolation_config_ds. where ^missing(product_line_group);
  quit;


  data extrapolation_report_md1(drop=ii hist_season);
    set &extrapolation_config_ds.(obs=1);
    end_season_week=input(scan(seasonality,2,"-"), 8.);
    extrapolation_season=hist_season;
    if end_season_week<52 then do;
      end_season_year=hist_season+1;
    end; else do;
      end_season_year=hist_season;
    end;
    do ii=1 to 52;
      mid_season_week=ii;
      if mid_season_week > end_season_week then do;
        mid_season_year=end_season_year-1;
      end; else do;
        mid_season_year=end_season_year;
      end;
      if week=mid_season_week or missing(week) then output;
    end;
  run;

  proc sort data=extrapolation_report_md1 out=extrapolation_report_wks(keep=mid_season_year mid_season_week);
    by mid_season_year mid_season_week;
  run;

  proc sort data=extrapolation_report_md1 out=extrapolation_report_wk_end(keep=end_season_year end_season_week) nodupkey;
    by end_season_year end_season_week;
  run;

  data extrapolation_report_wk_all;
    set 
      extrapolation_report_wks(rename=(mid_season_year=year mid_season_week=week)) 
      extrapolation_report_wk_end(rename=(end_season_year=year end_season_week=week));
  run;

  proc sort data=extrapolation_report_wk_all nodupkey;
    by year week;
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
  run;

  proc sql;
    create table filelist3 as
    select a.year, a.week, b.dsname from extrapolation_report_wk_all a
    left join filelist2 b on a.year=b.year and a.week=b.week
    where ^missing(b.dsname);
  quit;

  data filelist4;
    set filelist3;
    order=_n_;
  run;

  proc sql noprint;
  select count(*) into : flcnt from filelist4;
  quit;
  
  %do i=1 %to &flcnt.;
    proc sql noprint;
      select dsname, year, week into :dsname trimmed, :year trimmed, :week trimmed from filelist4 where order = &i.;
    quit;
/*trinh*/
    data sales_history_tmp1(drop=sls_off shipto_cntry order_type rsn_rej_cd);
      set shw.&dsname. (keep=sls_org soldto_nr sls_off shipto_cntry material variety SchedLine_Cnf_deldte cnf_qty order_type mat_div rsn_rej_cd);
      length unique_code $10. year 8. week 8.;
      unique_code=cats(sls_org, sls_off, shipto_cntry);
      year=&year.;
      week=&week.;
      if mat_div in (&mat_div_quote.) and ((mat_div in ('6B', '6C') and order_type in ('ZYPD', 'ZFD1', 'ZYPL', 'ZMTO')) or (mat_div='6A' and order_type in ('YQOR', 'ZMTO'))) and missing(rsn_rej_cd) and ^missing(SchedLine_Cnf_deldte) then output;
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

    %if &i.=1 %then %do;
      data sales_history;
        set sales_history_tmp1;
      run;
    %end; %else %do;
      data _null_;
        sleep=sleep(0);
      run;
      proc append base=sales_history data=sales_history_tmp1;
      run;
    %end;
  %end;
  
  data sales_history1(drop=rc);
    set sales_history;
    length region $3. territory $3. country $6. macro_region $3.;
    retain macro_region;
    if _n_=1 then do;
      declare hash cl(dataset: 'dmimport.Country_lookup');
        rc=cl.DefineKey ('unique_code');
        rc=cl.DefineData ('region', 'territory', 'country');
        rc=cl.DefineDone();
      declare hash stl(dataset: 'dmimport.Soldto_nr_lookup');
        rc=stl.DefineKey ('soldto_nr');
        rc=stl.DefineData ('region', 'territory', 'country');
        rc=stl.DefineDone();
        macro_region="&region.";
    end;

    rc=cl.find(); /*gets territory and country from country_lookup*/
    if region='BI' then do;
      rc=stl.find();/*gets territory and country from soldto_nr_lookup (if found overwrite the country_lookup)*/
    end;
    if region=macro_region and mat_div in (&mat_div_quote.) then output;
  run;

  data sales_history2(drop=rc);
    set sales_history1;
    length sub_unit 8.;

    if _n_=1 then do;
      declare hash material_assortment(dataset: 'dmproc.material_assortment');
        rc=material_assortment.DefineKey ('region', 'material');
        rc=material_assortment.DefineData ('sub_unit');
        rc=material_assortment.DefineDone(); 
    end;

    rc=material_assortment.find(); 

    if ^missing(sub_unit) then do;
      historical_sales=cnf_qty * sub_unit;
    end;

  run;

  data sales_history3(keep=sls_org region year week species variety order_season country historical_sales product_line_group);
    set sales_history2;
    length  season_week_start season_week_end 
             Order_season_start order_year order_season order_week Order_yweek order_month macro_extrapolation_season 8.;
    length species $29. product_line_group $20. macro_product_line_group $20.;
    retain macro_extrapolation_season macro_product_line_group;
    if _n_=1 then do;
      declare hash pmd_assortment(dataset: 'dmproc.PMD_assortment');
        rc=pmd_assortment.DefineKey ('region', 'variety');
        rc=pmd_assortment.DefineData ('season_week_start', 'season_week_end', 'species', 'product_line_group');
        rc=pmd_assortment.DefineDone();
      macro_extrapolation_season=&extrapolation_season.;
      macro_product_line_group="&product_line_group";
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

    if order_season=macro_extrapolation_season and product_line_group=macro_product_line_group then output;
  run;

  
  %filter_orders(in_table=sales_history3, out_table=sales_history4);

  proc sql;
    create table sales_history_aggr as
    select region, year, week, species, variety, order_season, country, round(sum(historical_sales),1) as country_sales_aggr 
      from sales_history4 
      group by region, year, week, species, variety, order_season, country;
  quit;

  proc sql;
    create table country_mid_season_aggr as
    select b.* from extrapolation_report_wks a
    left join sales_history_aggr b on a.mid_season_year=b.year and a.mid_season_week=b.week;
  quit;

  proc sql;
    create table species_mid_season_aggr as
    select region, year, week, species, variety, order_season, round(sum(country_sales_aggr),1) as species_sales_aggr from country_mid_season_aggr a
    group by region, year, week, species, variety, order_season;
  quit;

  proc sql;
    create table country_end_season_aggr as
    select b.* from extrapolation_report_wk_end a
    left join sales_history_aggr b on a.end_season_year=b.year and a.end_season_week=b.week;
  quit;

  proc sql;
    create table species_end_season_aggr as
    select region, year, week, species, variety, order_season, round(sum(country_sales_aggr),1) as species_sales_aggr from country_end_season_aggr a
    group by region, year, week, species, variety, order_season;
  quit;

  proc sql;
    create table all_country_weeks as
    select b.*, a.* from 
      (select distinct mid_season_year, mid_season_week from extrapolation_report_md1) a
      inner join (select distinct species, variety, country from country_end_season_aggr) b on 1=1;
  quit;

  proc sql;
    create table all_species_weeks as
    select b.*, a.* from  
      (select distinct mid_season_year, mid_season_week from extrapolation_report_md1) a
      inner join (select distinct species, variety from species_end_season_aggr) b on 1=1;
  quit;

  proc sql;
  create table country_sales_percentage as
    select   e.region, 
            "&product_line_group." as product_line_group, 
            e.species, 
			e.variety, 
            "&mat_div." as mat_div,
            e.order_season as hist_season,
            a.country, 
            a.mid_season_year as mid_year,
            a.mid_season_week as mid_week,
            m.country_sales_aggr as sales, 
            e.country_sales_aggr as end_season_sales,  
            m.country_sales_aggr/e.country_sales_aggr as extrapolation_rate 
    from all_country_weeks a
    left join country_end_season_aggr e on a.species=e.species and a.variety=e.variety and a.country=e.country  
    left join country_mid_season_aggr m on m.species=a.species and m.variety=e.variety and m.country=a.country and m.year=a.mid_season_year and m.week=a.mid_season_week;
  quit;

  proc sql;
  create table species_sales_percentage as
    select   e.region, 
            "&product_line_group." as product_line_group, 
            e.species,
			e.variety, 
            "&mat_div." as mat_div,
            e.order_season as hist_season, 
            a.mid_season_year as mid_year,
            a.mid_season_week as mid_week,
            m.species_sales_aggr as sales, 
            e.species_sales_aggr as end_season_sales,  
            m.species_sales_aggr/e.species_sales_aggr as extrapolation_rate 
    from all_species_weeks a
    left join species_end_season_aggr e on a.species=e.species and a.variety=e.variety
    left join species_mid_season_aggr m on m.species=a.species and m.variety=e.variety and m.year=a.mid_season_year and m.week=a.mid_season_week;
  quit;

%mend extrapolation_extraction;