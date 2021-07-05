/***********************************************************************/
/*Type: Report*/
/*Use: Fill in parameters in metadata.xlsx, sheet=s961 and press run*/
/*Purpose: Create forecast report with 2 seperate steps*/
/*OUT: Replacer list excel report and ZDEMAND format upload file in in upload_report_folder (check configuration.sas)*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\cleanup_xlsx_bak_folder.sas";
%include "&sas_applications_folder.\read_metadata.sas";
%include "&sas_applications_folder.\filter_orders.sas";


%macro s961_upload_preparation(forecast_report_file=,
                                            region=,
                                            material_division=,
                                            season=,
                                            historical_season=,
                                            product_line_group=
                                            );

  PROC IMPORT OUT=forecast_report_file 
              DATAFILE="&forecast_report_file."
              DBMS=  EXCELCS  REPLACE;
              RANGE="Variety level fcst$A8:AX"; 
  RUN;

  proc contents data=forecast_report_file out=forecast_contents noprint;
  run;

  data forecast_cols;
    length varnum 8. columnname $32.;
    varnum=3; /*Excel column C*/
    columnname="Country"; 
    output;
    varnum=10; /*Excel column J*/
    columnname="Variety";
    output;
    varnum=35; /*Excel column AI*/
    columnname="Netproposal0";
    output;
    varnum=41; /*Excel column AO*/
    columnname="Netproposal1";
    output;
    varnum=47; /*Excel column AU*/
    columnname="Netproposal2";
    output;
  run;

  proc sql noprint;
    select compress(name||'=_'||columnname) into :renamestring separated by ' ' from forecast_cols fcols
    left join forecast_contents frcst on fcols.varnum=frcst.varnum;
  quit;

  data forecast_report (keep=variety country total_demand Netproposal0 Netproposal1 Netproposal2);
    set forecast_report_file(rename=(&renamestring.));
    length country $6. variety Netproposal0 Netproposal1 Netproposal2 8.;
    country=strip(_country);
    variety=input(strip(_variety), 8.);
    Netproposal0=input(_Netproposal0, comma20.);
    Netproposal1=input(_Netproposal1, comma20.);
    Netproposal2=input(_Netproposal2, comma20.);
    if missing(Netproposal0) then Netproposal0=0;
    if missing(Netproposal1) then Netproposal1=0;
    if missing(Netproposal2) then Netproposal2=0;
    total_demand=Netproposal0+Netproposal1+Netproposal2;
    if ^missing(Variety) and country = "&region." then output;
  run;

  proc sql;
    create table bulk_var_mat as 
    select a.variety, c.variety_name, a.total_demand, c.current_plc as variety_plc, b.material, b.fps_material_name as material_description, b.process_stage as Process_stage
    from forecast_report a
    left join dmproc.material_assortment b on a.country=b.region and a.variety=b.variety and b.bulk_6a=1 and b.material_plc_current^='G2' and material_division='6A'
    left join dmproc.pmd_assortment c on a.country=c.region and a.variety=c.variety;
  quit;

  proc sql;
    create table varieties_not_in_pmd as
    select distinct variety from forecast_report where variety not in (select variety from dmproc.pmd_assortment where region="&region.");
  quit;

  proc sql;
  create table forecast_report0 as
    select * from forecast_report 
    where variety in (select variety from dmproc.material_assortment where bulk_6a=1 and material_plc_current^='G2' and material_division='6A' and region="&region.")
    and variety not in (select variety from varieties_not_in_pmd);
  quit;

  proc sql;
    create table forecast_report1 as
    select a.variety, a.country as region, a.total_demand, a.Netproposal0, a.Netproposal1, a.Netproposal2, b.material, b.material_division as mat_div, &historical_season. as order_season, b.process_stage as process_stage  
    from forecast_report0 a
    left join dmproc.material_assortment b on a.country=b.region and a.variety=b.variety and b.bulk_6a=1 and b.material_plc_current^='G2' and b.material_division='6A';
  quit;

  %filter_orders(in_table=dmproc.orders_all, out_table=orders_filtered);
/*Call order report*/
  data orders1(keep=material
                    variety
                    mat_div
                    region
                    sub_unit
                    actual_sales
                    order_year
                    order_month_season
                    order_month
                    process_stage);
    set orders_filtered;
    if mat_div in ("&material_division.") and region="&region." and order_month_season=&historical_season. then output;
  run;

/*Filter orders1 with forecast materials*/
  proc sql;
    create table orders2 as
    select   f.region, f.variety, f.process_stage, 
             c.material, c.material_division as mat_div, c.sub_unit, 
            coalesce(o.actual_sales, 0) as actual_sales, 
            coalesce(o.order_year, &historical_season.) as order_year, 
            coalesce(o.order_month_season, &historical_season.) as order_month_season, 
            coalesce(o.order_month, 12) as order_month
    from forecast_report1 f
    left join orders1 o on o.region=f.region and o.variety=f.variety and f.process_stage=o.process_stage
    left join dmproc.material_assortment c on f.region=c.region and f.variety=c.variety and f.process_stage=c.process_stage and c.bulk_6a=1;
  quit;

/*Filter-out missing sub_unit and missing order_year*/
  proc sql;
    create table orders3 as
    select * from orders2
    where ^missing(sub_unit) and ^missing(order_year);
  quit;

/*Add product information from FPS and PMD*/
  proc sql;
    create table orders4 as
    select a.*, b.species, b.series, b.variety_name, b.current_plc as var_current_plc, c.fps_material_name as material_name from orders3 a
      left join dmproc.PMD_Assortment b on a.region=b.region and a.variety=b.variety
      left join dmproc.material_assortment c on a.region=c.region and a.variety=c.variety and a.material=c.material;
  quit;

/*Calculate sales per series for var plc F2 and F3*/
  proc sql;
    create table orders_average_series as
    select region, order_month_season, series, process_stage, sum(actual_sales) as sum_series from orders4
    where var_current_plc in ('F2', 'F3')
      group by region, order_month_season, series, process_stage;
  quit;

  /*Calculate sales per species for var plc F2 and F3*/
  proc sql;
    create table orders_average_species as
    select region, order_month_season, species, process_stage, sum(actual_sales) as sum_species from orders4
    where var_current_plc in ('F2', 'F3')
      group by region, order_month_season, species, process_stage;
  quit;

  /*Calculate sales per series per month for var plc F2 and F3*/
  proc sql;
    create table orders_average_series_month as
    select region, order_month_season, series, process_stage, order_year, order_month, sum(actual_sales) as sum_series_month 
    from orders4 
    where var_current_plc in ('F2', 'F3')
    group by region, order_month_season, series, process_stage, order_year, order_month;
  quit;

  /*Calculate sales per species per month for var plc F2 and F3*/
  proc sql;
    create table orders_average_species_month as
    select region, order_month_season, species, process_stage, order_year, order_month, sum(actual_sales) as sum_species_month 
    from orders4 
    where var_current_plc in ('F2', 'F3')
    group by region, order_month_season, species, process_stage, order_year, order_month;
  quit;

  /*Calculate series month percentage*/
  proc sql;
    create table series_per_month as
    select a.region, a.order_month_season, a.series, a.process_stage, a.order_year, a.order_month, 
      a.sum_series_month, b.sum_series, put(order_year, 4.)||put(order_month, z2.) as yearmonth, 
      coalesce((a.sum_series_month/b.sum_series),0) as series_month_percentage 
      from orders_average_series_month a
    left join orders_average_series b on a.region=b.region and a.order_month_season=b.order_month_season and a.series=b.series and a.process_stage=b.process_stage;
  quit;

  /*Calculate species month percentage*/
  proc sql;
    create table species_per_month as
    select a.region, a.order_month_season, a.species, a.process_stage, a.order_year, a.order_month, 
      a.sum_species_month, b.sum_species, put(order_year, 4.)||put(order_month, z2.) as yearmonth, 
      coalesce((a.sum_species_month/b.sum_species),0) as species_month_percentage 
      from orders_average_species_month a
    left join orders_average_species b on a.region=b.region and a.order_month_season=b.order_month_season and a.species=b.species and a.process_stage=b.process_stage;
  quit;

  /*Transpose data*/
  proc transpose data=series_per_month out=series_per_month1(drop=_name_) prefix=M_;
    by region series process_stage;
    id yearmonth;
    var series_month_percentage;
  run;

  /*Transpose data*/
  proc transpose data=species_per_month out=species_per_month1(drop=_name_) prefix=M_;
    by region species process_stage;
    id yearmonth;
    var species_month_percentage;
  run;

/*Calculate sum_variety from order data*/
  proc sql;
    create table orders_aggr_season as
    select region, order_month_season, variety, sum(actual_sales) as sum_variety
    from orders3 
    group by region, order_month_season, variety;
  quit;

/*Calculate sum_ps from order data*/
  proc sql;
    create table orders_aggr_ps as
    select region, order_month_season, variety, process_stage, sum(actual_sales) as sum_ps 
    from orders3 
    group by region, order_month_season, variety, process_stage;
  quit;

/*Calculate ps_percentage*/
  proc sql;
    create table orders_per_ps as
    select w.region, w.order_month_season, w.variety, w.process_stage, w.sum_ps, s.sum_variety, 
          w.sum_ps/s.sum_variety as ps_percentage, . as adjusted_percentage  
    from orders_aggr_ps w
    left join orders_aggr_season s on w.region=s.region and w.variety=s.variety and w.order_month_season=s.order_month_season
    order by w.region, w.order_month_season, w.variety, w.process_stage;
  quit;

/*Add information from FPS and PMD*/
  proc sql;
    create table orders_per_ps1 as
    select a.*, b.species, b.series, b.variety_name, c.fps_material_name as material_name, c.material, b.current_plc as variety_plc, c.material_plc_current as material_plc
    from orders_per_ps a
    left join dmproc.PMD_Assortment b on a.region=b.region and a.variety=b.variety
    left join dmproc.material_assortment c on a.region=c.region and a.variety=c.variety and a.process_stage=c.process_stage and c.bulk_6a=1 and c.material_plc_current^='G2' and material_division='6A'
    order by a.variety, c.material;
  quit;

 /*Calculate total orders per month*/
  proc sql;
    create table orders_aggr_month as
    select region, order_month_season, variety, process_stage, order_year, order_month, sum(actual_sales) as sum_month 
    from orders3 
    group by region, order_month_season, variety, process_stage, order_year, order_month;
  quit;

 /*Calculate monthly percentage*/
  proc sql;
    create table orders_per_month as
    select w.region, w.order_month_season, w.variety, w.process_stage, w.order_year, w.order_month, w.sum_month, s.sum_ps, w.sum_month/s.sum_ps as month_percentage 
    from orders_aggr_month w
    left join orders_aggr_ps s on w.region=s.region and w.variety=s.variety and w.order_month_season=s.order_month_season and w.process_stage=s.process_stage
    order by w.region, w.order_month_season, w.variety, w.process_stage, w.order_year, w.order_month;
  quit;

  %let season_start_week=%scan(&seasonality.,1,'-');
  data all_months (keep=order_year order_month);
    length hist_year order_year order_month order_season_start 8.;
    format order_season_start yymmdd10.;

    hist_year=&historical_season.;
    season_week_start=&season_start_week.;
    order_season_start=input(put(hist_year, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
    order_season_start_ym=input(put(year(order_season_start),4.)||put(month(order_season_start),z2.), 6.);

    do i= 1 to 12;
      i_ym=put(hist_year, 4.)||put(i, z2.);
      order_month=i;
      if order_season_start_ym < i_ym or (order_season_start_ym=i_ym and day(order_season_start)<=15) then do;
          order_year=&historical_season.;
      end; else do;
          order_year=&historical_season.+1;
      end;
      output;
    end;
  run;

  proc sql;
    create table all_months_combination as
    select * from all_months a
    join (select distinct region, variety, process_stage from forecast_report1) on 1=1
    order by region, variety, process_stage, order_year, order_month;
  quit;

  proc sql;
    create table all_months_combination1 as
    select a.*, c.material from all_months_combination a
    left join dmproc.material_assortment c on a.region=c.region and a.variety=c.variety and a.process_stage=c.process_stage and c.bulk_6a=1 and c.material_plc_current^='G2' and material_division='6A';
  quit;

%mend s961_upload_preparation;

%macro s961_create_upload_mat_perc(forecast_report_file=,
                                            region=,
                                            material_division=,
                                            season=,
                                            historical_season=,
                                            product_line_group=
                                            );

  %s961_upload_preparation(forecast_report_file=%quote(&forecast_report_file.),
                                            region=&region.,
                                            material_division=&material_division.,
                                            season=&season.,
                                            historical_season=&historical_season.,
                                            product_line_group=&product_line_group.);

  proc sql;
    create table orders_per_month1 as
    select a.region, a.variety, a.process_stage, a.order_year, a.order_month, 
            coalesce(b.order_month_season, &historical_season.) as order_month_season, 
            coalesce(b.sum_month, 0) as sum_month,
            (select distinct c.sum_ps from orders_per_month c where a.variety=c.variety and a.process_stage=c.process_stage and ^missing(c.sum_ps)) as sum_ps,
            coalesce(b.month_percentage, 0) as month_percentage,
            put(a.order_month, z2.) as yearmonth 
    from all_months_combination a
    left join orders_per_month b on a.region=b.region and a.variety=b.variety and a.process_stage=b.process_stage and a.order_year=b.order_year and a.order_month=b.order_month;
  quit;

  proc transpose data=orders_per_month1 out=orders_per_month2(drop=_name_) prefix=M_;
    by region variety process_stage;
    id yearmonth;
    var month_percentage;
  run;

  data orders_per_month3;
    length sum 8.;
    set orders_per_month2;
    sum=sum(of M_:);
  run;

  proc sql;
    create table order_per_month4 as
    select b.species, b.series, b.variety_name, c.fps_material_name as material_name, c.material, 
          c.replaced_by as replaced_by_suggestion, b.current_plc as variety_plc, c.material_plc_current as material_plc,
          case when sum=0 then 'NO_HISTORY' end as REPLACE, a.* 
    from orders_per_month3 a
    left join dmproc.PMD_Assortment b on a.region=b.region and a.variety=b.variety
    left join dmproc.material_assortment c on a.region=c.region and a.variety=c.variety and a.process_stage=c.process_stage and c.bulk_6a=1 and c.material_plc_current^='G2' and material_division='6A'
    order by a.variety, c.material;
  quit;

  proc sql noprint; 
    select replacer_path_file, dir_lvl31 into :replacer_path_file trimmed, :cleanup_folder trimmed from upload_folders;
  quit;

  x "del &replacer_path_file.";

  proc sql;
    create table adjust as
    select a.*, b.var_count, c.total_demand from orders_per_ps1 a
    left join (select variety, count(*) as var_count from orders_per_ps1 group by variety) b on a.variety=b.variety
    left join forecast_report c on a.variety=c.variety;
  quit;

  quit;
  data adjust1;
    retain region  order_month_season  variety  var_count  variety_plc  material  Process_stage  material_plc  sum_ps  sum_variety total_demand  ps_percentage  adjusted_percentage  Species  Series  Variety_name  material_name;
    set adjust;
  run;

  proc export 
    data=adjust1 
    dbms=xlsx 
    outfile="&replacer_path_file." replace;
    sheet="adjust";
  run;


  proc sql;
    create table replace as
    select a.*, b.total_demand from order_per_month4 a
    left join forecast_report b on a.variety=b.variety;
  quit;

  data series_per_month2;
    set series_per_month1;
    sum=sum(of m:);
    if sum^=0 then output;
  run;

  data species_per_month2;
    set species_per_month1;
    sum=sum(of m:);
    if sum^=0 then output;
  run;

  data series_averages_info;
    retain region series process_stage sum;
    set series_per_month2;
  run;

  data species_averages_info;
    retain region species process_stage sum;
    set species_per_month2;
  run;

  proc sql;
    create table replace1 as
    select 
      a.*, 
      b.sum as series_average_sum,
      c.sum as species_average_sum 
    from replace a
    left join series_averages_info b on a.series=b.series and a.process_stage=b.process_stage
    left join species_averages_info c on a.species=c.species and a.process_stage=c.process_stage;
  quit;

  data replace2;
    retain region  Species  Series  variety  Variety_name  variety_plc total_demand  material  material_name  Process_stage  material_plc  replaced_by_suggestion  REPLACE series_average_sum species_average_sum sum;
    set replace1;
  run;

  proc export 
    data=replace2
    dbms=xlsx 
    outfile="&replacer_path_file.";
    sheet="replace";
  run;

  proc export 
    data=series_averages_info
    dbms=xlsx 
    outfile="&replacer_path_file.";
    sheet="series_averages_info";
  run;

  proc export 
    data=species_averages_info
    dbms=xlsx 
    outfile="&replacer_path_file.";
    sheet="species_averages_info";
  run;

  proc export 
    data=bulk_var_mat
    dbms=xlsx 
    outfile="&replacer_path_file.";
    sheet="variety_bulk_material";
  run;

  proc export 
    data=varieties_not_in_pmd
    dbms=xlsx 
    outfile="&replacer_path_file.";
    sheet="varieties_not_in_pmd";
  run;

  %cleanup_xlsx_bak_folder(cleanup_folder=%str(&cleanup_folder.));

%mend s961_create_upload_mat_perc;

%macro s961_generate_upload(forecast_report_file=,
                      region=,
                      material_division=,
                      season=,
                      historical_season=,
                      product_line_group=
                      );

  %s961_upload_preparation(forecast_report_file=%quote(&forecast_report_file.),
                                            region=&region.,
                                            material_division=&material_division.,
                                            season=&season.,
                                            historical_season=&historical_season.,
                                            product_line_group=&product_line_group.);
  %let season0=&season.;
  %let season1=%eval(&season.+1);
  %let season2=%eval(&season.+2);

  PROC IMPORT OUT=mat_percentage_adjusted
    DATAFILE="&Split_configuration_file."
    DBMS=  XLSX  REPLACE;
    sheet="adjust";
  RUN;

  data mat_percentage_adjusted1(drop=_:);
    set mat_percentage_adjusted(rename=(adjusted_percentage=_adjusted_percentage));
    length adjusted_percentage 8.;
    adjusted_percentage=input(_adjusted_percentage, comma20.);
  run;

  proc sql;
    create table mat_percentage_dist1 as
    select d.*, p.mat_dist_percentage 
    from mat_percentage_adjusted1 d 
    left join (select e.*, (e.ps_percentage/s.sum_var) as mat_dist_percentage from mat_percentage_adjusted1 e 
              left join (select variety, sum(ps_percentage) as sum_var from mat_percentage_adjusted1 where missing(adjusted_percentage) group by variety) s on e.variety=s.variety
              where missing(e.adjusted_percentage)) p on d.material=p.material;
  quit;

  proc sql;
    create table mat_adjusted_to_dist as 
    select variety, sum(ps_percentage-adjusted_percentage) as perc_diff from mat_percentage_adjusted1 where ^missing(adjusted_percentage) group by variety;
  quit;

  proc sql;
    create table mat_perc_distributed as
    select region, variety, material, coalesce(distributed_mat_perc, adjusted_percentage, ps_percentage) as final_mat_percentage from 
    (select a.*, b.perc_diff, a.ps_percentage+(b.perc_diff*a.mat_dist_percentage) as distributed_mat_perc from mat_percentage_dist1 a
    left join mat_adjusted_to_dist b on a.variety=b.variety and missing(a.adjusted_percentage));
  quit;

  proc sql;
    create table mat_net_amounts as
    select a.*, (b.netproposal0*a.final_mat_percentage) as mat_proposal0,
                (b.netproposal1*a.final_mat_percentage) as mat_proposal1,
                (b.netproposal2*a.final_mat_percentage) as mat_proposal2
    from mat_perc_distributed a 
    left join forecast_report1 b on a.variety=b.variety and a.material=b.material;
  quit;

  data mat_net_amounts1(drop=mat_proposal0 mat_proposal1 mat_proposal2);
    length season 8.;
    set mat_net_amounts;
    season=&season0.;
    mat_proposal=round(mat_proposal0,1);
    output;
    season=&season1.;
    mat_proposal=round(mat_proposal1,1);
    output;
    season=&season2.;
    mat_proposal=round(mat_proposal2,1);
    output;
  run;

  PROC IMPORT OUT=mat_replace_raw
    DATAFILE="&Split_configuration_file."
    DBMS=  XLSX  REPLACE;
    sheet="replace";
  RUN;

  data 
    mat_replace(keep=region variety material replace_material) 
    series_mat_average(keep=region variety material) 
    species_mat_average(keep=region variety material)
    mat_noreplace(keep=region variety material);
    set mat_replace_raw(rename=(replace=_replace));
    if ^missing(_replace) then do;
      if strip(_replace)="SERIES" then do;
        replace=_replace;
        output series_mat_average;
      end; else if strip(_replace)="SPECIES" then do;
        replace=_replace;
        output species_mat_average;
      end; else do;
        replace_material=input(strip(_replace),8.);
        output mat_replace;
      end;
    end; else do;
      output mat_noreplace;
    end;
  run;

  proc sql;
    create table orders_per_month_mat as
    select a.*, b.material from orders_per_month a
    left join dmproc.material_assortment b on a.variety=b.variety and a.process_stage=b.process_stage and b.bulk_6a=1 and b.material_plc_current^='G2' and material_division='6A';
  quit;

  proc sql;
    create table orders_per_month_norep as
    select b.* from mat_noreplace a
    left join orders_per_month_mat b on a.material=b.material;
  quit;

  proc sql;
    create table orders_per_month_rep as
    select a.region, a.variety, a.material, b.order_month_season, c.process_stage as process_stage, b.order_year, b.order_month, b.sum_month, b.sum_ps, b.month_percentage from mat_replace a
    left join orders_per_month_mat b on a.replace_material=b.material
    left join dmproc.material_assortment c on a.region=c.region and a.variety=c.variety and a.material=c.material;
  quit;

  proc sql;
    create table series_mat_average1 as 
    select a.variety, a.material, b.series, b.species, c.process_stage as process_stage from series_mat_average a
      left join dmproc.PMD_Assortment b on a.region=b.region and a.variety=b.variety
      left join dmproc.material_assortment c on a.region=c.region and a.variety=c.variety and a.material=c.material;
  quit;

  proc sql;
    create table species_mat_average1 as 
    select a.variety, a.material, b.species, b.species, c.process_stage as process_stage from species_mat_average a
      left join dmproc.PMD_Assortment b on a.region=b.region and a.variety=b.variety
      left join dmproc.material_assortment c on a.region=c.region and a.variety=c.variety and a.material=c.material;
  quit;


/*  proc sql;*/
/*    create table orders_per_month_average as*/
/*    select a.variety, a.material, b.* from mat_average1 a*/
/*    left join series_per_month b on a.series=b.series and a.process_stage=b.process_stage;*/
/*  quit;*/

  proc sql;
    create table orders_per_month_avg_series as
    select a.variety, a.material, b.* from series_mat_average1 a
    left join series_per_month b on a.series=b.series and a.process_stage=b.process_stage;
  quit;

  proc sql;
    create table orders_per_month_avg_species as
    select a.variety, a.material, b.* from species_mat_average1 a
    left join species_per_month b on a.species=b.species and a.process_stage=b.process_stage;
  quit;

  data orders_per_month1;
    set orders_per_month_norep
        orders_per_month_rep
        orders_per_month_avg_series(drop=series process_stage yearmonth rename=(sum_series_month=sum_month sum_series=sum_ps series_month_percentage=month_percentage))
        orders_per_month_avg_species(drop=species process_stage yearmonth rename=(sum_species_month=sum_month sum_species=sum_ps species_month_percentage=month_percentage));
  run;

  data final_distribution;
    set orders_per_month1(rename=(month_percentage=total_percentage));
  run;

  proc sql;
    create table final_distribution1 as
    select a.region, a.variety, a.material, a.order_year, a.order_month, coalesce(b.total_percentage,0) as total_percentage
    from all_months_combination1 a
    left join final_distribution b on a.region=b.region and a.variety=b.variety and a.material=b.material and a.order_year=b.order_year and a.order_month=b.order_month;
  quit;

  proc sql;
    create table final_distribution2 as
    select a.*, b.final_mat_percentage, b.mat_proposal0, b.mat_proposal1, b.mat_proposal2, c.fps_material_name as material_basic_description_en, c.sub_unit, c.process_stage as process_stage
    from final_distribution1 a
    left join mat_net_amounts b on a.region=b.region and a.variety=b.variety and a.material=b.material
    left join dmproc.material_assortment c on a.region=c.region and a.variety=c.variety and a.material=c.material;
  quit;

  proc sql;
    create table final_distribution3 as
    select region, variety, process_stage, order_year, order_month, sum(mat_proposal0*total_percentage) as ps0 , sum(mat_proposal1*total_percentage) as ps1, sum(mat_proposal2*total_percentage) as ps2
    from final_distribution2 
    group by region, variety, process_stage, order_year, order_month;
  quit;

  data final_distribution4;
    set final_distribution3(rename=(order_year=_order_year));
    Confirmed_Sales_Forecast=round(ps0,1);
    order_year=_order_year+&year_offset.;
      output;
    Confirmed_Sales_Forecast=round(ps1,1);
    order_year=_order_year+&year_offset.+1;
      output;
    Confirmed_Sales_Forecast=round(ps2,1);
    order_year=_order_year+&year_offset.+2;
      output;
  run;

  proc sql;
    create table final_distribution5 as
    select a.*, b.material, b.fps_material_name as material_description 
    from final_distribution4 a
    left join dmproc.material_assortment b on a.region=b.region and a.variety=b.variety and a.process_stage=b.process_stage and b.bulk_6a=1 and b.material_plc_current^='G2' and b.material_division='6A';
  quit;

  proc sql;
    create table final_distribution6 as
    select a.*, b.current_plc as variety_plc
    from final_distribution5 a
    left join dmproc.pmd_assortment b on a.region=b.region and a.variety=b.variety
    where variety_plc ^='G2' and ^missing(Confirmed_Sales_Forecast) and Confirmed_Sales_Forecast>0;
  quit;

  proc sql;
    create table final_distribution7 as
    select 'S961' as Info_Str,
            put(order_year,4.)||put(order_month,z2.) as Month, 
            0 as Week, 
            0 as Period, 
            'ZF' as Div,
             'NL02' as Sales_Org,
            'NL50' as Sales_Off,
            'NL03' as Plant,
            variety as Variety,
            material as Material,
            Material_Description,
            '+++' as Sal, 
            '++++++++++' as SRep,
            '++++++++++' as Sold_to_Party,
            Confirmed_Sales_Forecast,
            0 as Returns_Planned,
            0 as Market_Uncertainity,
            Confirmed_Sales_Forecast as Confirmed_Sales_Plan,
            'KS' as Base_Unit 
    from final_distribution6 
    order by variety, material, month;
  quit;

  proc sql noprint; 
    select s961_upload_path_file into :s961_upload_path_file trimmed from upload_folders;
  quit;

  PROC EXPORT DATA=final_distribution7
   OUTFILE="&s961_upload_path_file."
       DBMS=TAB REPLACE;
   PUTNAMES=YES;
  RUN;

%mend s961_generate_upload;

%macro upload_s961();

  %read_metadata(sheet=s961);

  data s961_md;
    set dmimport.s961_md;
    if step1='Y' or step2='Y' then output;
  run;

  proc sql noprint;
    select count(*) into :report_cnt from s961_md;
  quit;
  
  %do ii=1 %to &report_cnt.;

    data _null_;
      set s961_md;
      if _n_=&ii then do;
        call symput('Step1', strip(Step1));
        call symput('Forecast_report_file', strip(Forecast_report_file));
        call symput('Region', strip(Region));
        call symput('Mat_div', strip(tranwrd(Mat_div, ',','_')));
        call symput('Season', strip(Season));
        call symput('Historical_season', strip(Historical_season));
        call symput('Product_line_group', strip(Product_line_group));
        call symput('Seasonality', strip(Seasonality));
        call symput('Step2', strip(Step2));
        call symput('Split_configuration_file', strip(Split_configuration_file));
      end;
    run;

    %if "&step1."="Y" or "&step2."="Y" %then %do;

      data upload_folders(keep=dir_lvl31 replacer_path_file s961_upload_path_file);
        dir_lvl1="&upload_report_folder.";
        product_line_group="&product_line_group.";
        seasonality="&seasonality.";
        seasonality1=tranwrd(seasonality, '-', '_');
        Mat_div="&Mat_div.";
        region="&region.";
        dir1=catx('_',product_line_group, seasonality1, region, Mat_div);
        dir_lvl2=catx('\',dir_lvl1, dir1);
        dir2="&season.";
        dir_lvl3=catx('\',dir_lvl2, dir2);
        dir31=put(today(),yymmn.);
        dir_lvl31=catx('\',dir_lvl3, dir31);
        rc=dcreate(dir1,dir_lvl1);
        rc=dcreate(dir2,dir_lvl2);
        rc=dcreate(dir31,dir_lvl3);
        replacer_filename=catx('_', compress(put(today(),yymmdd10.),,'kd'), compress(put(time(), hhmm.),,'kd'), "replacer_list.xlsx");
        replacer_path_file=catx('\', dir_lvl31, replacer_filename);
        s961_upload_filename=cats( 'S961_', region, '_', product_line_group, '_', dir2, '_upload_',  compress(put(today(),yymmdd10.),,'kd'), '_', compress(put(time(), hhmm.),,'kd'), ".txt");
        s961_upload_path_file=catx('\', dir_lvl31, s961_upload_filename);
      run;

      %let year_offset=%eval(&season.-&historical_season.);
    %end;

    %if "&step1."="Y" %then %do;  
      %s961_create_upload_mat_perc(forecast_report_file=%quote(&forecast_report_file.),
                                            region=&region.,
                                            material_division=&Mat_div.,
                                            season=&season.,
                                            historical_season=&historical_season.,
                                            product_line_group=&product_line_group.);
    %end;

    %if "&step2."="Y" %then %do;
      %s961_generate_upload(forecast_report_file=%quote(&forecast_report_file.),
                                            region=&region.,
                                            material_division=&Mat_div.,
                                            season=&season.,
                                            historical_season=&historical_season.,
                                            product_line_group=&product_line_group.);
    %end;
  %end;

%mend upload_s961;

%upload_s961();
