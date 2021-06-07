/***********************************************************************/
/*Type: Report*/
/*Use: Fill in parameters in metadata.xlsx, sheet=s960 and press run*/
/*Purpose: Create forecast report with 1 step*/
/*OUT: Summary excel report and ZDEMAND format upload file in upload_report_folder (check configuration.sas)*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\cleanup_xlsx_bak_folder.sas";
%include "&sas_applications_folder.\read_metadata.sas";

%macro s960_upload_preparation(forecast_report_file=,
                                            region=,
                                            mat_div=,
                                            season=,
                                            product_line_group=,
                                            seasonality=  
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
    create table forecast_report1 as
    select b.product_line, b.species_code as species, a.variety, b.material_bulk as material, b.product_form, d.process_stage_percentage, a.total_demand
    from forecast_report a
    left join dmimport.BI_seed_assortment b on a.variety=b.variety
    left join dmimport.BI_process_stage_split d on upper(strip(b.product_line)) = upper(d.product_line) and b.species_code=d.species and b.product_form=d.process_stage
    order by b.product_line, b.species_code, a.variety, b.material_bulk, b.product_form;
  quit;

  proc sql;
    create table varieties_without_bulk_material as
    select distinct * from forecast_report1 a
    where missing(material);
  quit;

  proc sql;
    create table variety_ps_sum as
      select variety, sum(process_stage_percentage) as ps_sum from forecast_report1 group by variety;
  quit;

  proc sql;
    create table variety_ps_to_distribute as
      select variety, ps_sum, 1-ps_sum as ps_to_distribute from variety_ps_sum;
  quit;

  proc sql;
    create table variety_ps_count as
    select variety, count(*) as variety_ps_cnt from forecast_report1 group by variety;
  quit;

  proc sql;
    create table variety_redistribution as
      select a.*, coalesce(b.ps_to_distribute/c.variety_ps_cnt,0) as redistribution from forecast_report1 a 
      left join variety_ps_to_distribute b on a.variety=b.variety
      left join variety_ps_count c on a.variety=c.variety;
  quit;

  proc sql;
    create table forecast_report2 as 
      select a.product_line, a.species, a.variety, a.product_form, a.total_demand, a.material, a.process_stage_percentage, a.redistribution, 
              a.process_stage_percentage+a.redistribution as total_process_stage_percentage, b.month, b.month_percentage 
      from variety_redistribution a 
      left join dmimport.BI_seasonality b on upper(strip(a.product_line)) = upper(b.product_line) and a.species=b.species
      where ^missing(b.month)
      order by a.product_line, a.species, a.variety, a.material, a.product_form, a.process_stage_percentage, a.redistribution, total_process_stage_percentage,a.total_demand, b.month;
  quit;

  /*2020-06-07 - temporary deduplication, need to find the cause of the duplicates*/
  proc sort data=forecast_report2 nodupkey dupout=forecast_report2_dup;
  	by product_line species variety material product_form process_stage_percentage redistribution total_process_stage_percentage total_demand month;
  run;

  proc transpose data=forecast_report2 out=forecast_report3(drop=_name_) prefix=M_;
    by product_line species variety material product_form process_stage_percentage redistribution total_process_stage_percentage total_demand;
    id month;
    var month_percentage;
  run;

  proc sql;
    create table summary as
    select a.*, b.current_plc as variety_plc from forecast_report3 a
    left join dmproc.pmd_assortment b on a.variety=b.variety and b.region="&region."
    order by product_line, variety, product_form;
  quit;

  proc sql;
    create table summary1 as
    select a.*, b.var_count from summary a
    left join (select variety, count(*) as var_count from summary group by variety) b on a.variety=b.variety;
  quit;

  %let season_start_week=%scan(&seasonality.,1,'-');
  data all_months (keep=order_year order_month);
    length hist_year order_year order_month order_season_start 8.;
    format order_season_start yymmdd10.;

    hist_year=&season.;
    season_week_start=&season_start_week.;
    order_season_start=input(put(hist_year, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
    order_season_start_ym=input(put(year(order_season_start),4.)||put(month(order_season_start),z2.), 6.);

    do i= 1 to 12;
      i_ym=put(hist_year, 4.)||put(i, z2.);
      order_month=i;
      if order_season_start_ym < i_ym or (order_season_start_ym=i_ym and day(order_season_start)<=15) then do;
          order_year=&season.;
      end; else do;
          order_year=&season.+1;
      end;
      output;
    end;
  run;

  proc sql;
  create table forecast_report4 as
    select a.*, b.order_year as year from forecast_report2 a
    left join all_months b on a.month=b.order_month;
  quit;

%mend s960_upload_preparation;

%macro s960_create_upload_mat_perc(forecast_report_file=,
                                            region=,
                                            mat_div=,
                                            season=,
                                            product_line_group=,
                                            seasonality=
                                            );

  %s960_upload_preparation(forecast_report_file=%quote(&forecast_report_file.),
                                            region=&region.,
                                            mat_div=&mat_div.,
                                            season=&season.,
                                            product_line_group=&product_line_group.,
                                            seasonality=&seasonality.);

  proc sql noprint; 
    select summary_path_file, dir_lvl31 into :summary_path_file trimmed, :cleanup_folder trimmed from upload_folders;
  quit;

  x "del &summary_path_file.";

  data summary2;
    retain Product_line  species  variety  var_count  variety_plc  material  Product_Form  process_stage_percentage  redistribution  total_process_stage_percentage  total_demand  M_1  M_2  M_3  M_4  M_5  M_6  M_7  M_8  M_9  M_10  M_11  M_12;
    set summary1;
  run;

  proc export 
    data=summary2 
    dbms=xlsx 
    outfile="&summary_path_file." replace;
    sheet="summary";
  run;

  proc export 
    data=varieties_without_bulk_material 
    dbms=xlsx 
    outfile="&summary_path_file.";
    sheet="varieties_without_bulk_material";
  run;

  %cleanup_xlsx_bak_folder(cleanup_folder=%str(&cleanup_folder.));

%mend s960_create_upload_mat_perc;

%macro s960_generate_upload(forecast_report_file=,
                      region=,
                      mat_div=,
                      season=,
                      product_line_group=,
                      seasonality=
                      );

  %s960_upload_preparation(forecast_report_file=%quote(&forecast_report_file.),
                                            region=&region.,
                                            mat_div=&mat_div.,
                                            season=&season.,
                                            product_line_group=&product_line_group.,
                                            seasonality=&seasonality.);

  %let season0=&season.;
  %let season1=%eval(&season.+1);
  %let season2=%eval(&season.+2);

  proc sql;
    create table final_distribution1 as
    select a.variety, a.material, a.year, a.month, coalesce(a.total_process_stage_percentage,0)*coalesce(a.month_percentage,0) as total_percentage
    from forecast_report4 a
    left join forecast_report b on a.variety=b.variety
    order by a.variety, a.material, a.year, a.month;
  quit;

  proc sql;
    create table final_distribution2 as
    select a.*, b.Netproposal0*total_percentage as material_month_net0, b.Netproposal1*total_percentage as material_month_net1, b.Netproposal2*total_percentage as material_month_net2, d.bi_material_name
    from final_distribution1 a
    left join forecast_report b on a.variety=b.variety 
    left join dmimport.bi_seed_assortment c on a.variety=c.variety /*and ^missing(c.material_bulk)*/ and a.material=c.material_bulk
    left join dmproc.material_assortment d on c.material_bulk=d.material and d.region="&region.";
  quit;

  data final_distribution3(drop=material_month_net0 material_month_net1 material_month_net2);
    set final_distribution2(rename=(year=_year));
    Confirmed_Sales_Forecast=round(material_month_net0,1);
    year=_year;
      output;
    Confirmed_Sales_Forecast=round(material_month_net1,1);
    year=_year+1;
      output;
    Confirmed_Sales_Forecast=round(material_month_net2,1);
    year=_year+2;
      output;
  run;

  proc sql;
    create table final_distribution4 as
    select a.*, b.current_plc as variety_plc from final_distribution3 a
    left join dmproc.pmd_assortment b on a.variety=b.variety and b.region="&region.";
  quit;

  data final_distribution5;
    set final_distribution4;
    if variety_plc ^='G2' and Confirmed_Sales_Forecast>0 and ^missing(Confirmed_Sales_Forecast) then output;
  run;

  proc sql;
    create table final_distribution6 as
    select 'S960' as Info_Str,
            put(year,4.)||put(month,z2.) as Month, 
            0 as Week, 
            0 as Period, 
            'ZF' as Div,
            %if "&region."="BI" %then %do;
              'NL01' as Sales_Org,
              'NL81' as Sales_Off,
              'NL02' as Plant,
            %end;
            %if "&region."="FN" %then %do;
              'NL04' as Sales_Org,
              'NL84' as Sales_Off,
              'NL04' as Plant,
            %end;
            %if "&region."="JP" %then %do;
              'JP03' as Sales_Org,
              'JP02' as Sales_Off,
              'JP01' as Plant,
            %end;
            variety as Variety,
            material,
            bi_material_name as Material_Description,
            '+++' as Sales_Grp, 
            '++++++++++' as SRep,
            '++++++++++' as Sold_to_Party,
            Confirmed_Sales_Forecast,
            0 as Returns_Planned,
            0 as Market_Uncertainity,
            Confirmed_Sales_Forecast as Confirmed_Sales_Plan,
            'KS' as Base_Unit 
    from final_distribution5 a
    order by variety, material, month;
  quit;

  proc sql noprint; 
    select s960_upload_path_file into :s960_upload_path_file trimmed from upload_folders;
  quit;

  PROC EXPORT DATA=final_distribution6
   OUTFILE="&s960_upload_path_file."
       DBMS=TAB REPLACE;
   PUTNAMES=YES;
  RUN;

%mend s960_generate_upload;

%macro upload_s960();

  %read_metadata(sheet=s960);

  data s960_md;
    set dmimport.s960_md;
    if step1='Y' then output;
  run;

  proc sql noprint;
    select count(*) into :report_cnt from s960_md;
  quit;
  
  %do ii=1 %to &report_cnt.;

    data _null_;
      set s960_md;
      if _n_=&ii then do;
        call symput('Step1', strip(Step1));
        call symput('Forecast_report_file', strip(Forecast_report_file));
        call symput('Region', strip(Region));
        call symput('Mat_div', strip(tranwrd(Mat_div, ',','_')));
        call symput('Season', strip(Season));
        call symput('Product_line_group', strip(Product_line_group));
        call symput('Seasonality', strip(Seasonality));
      end;
    run;

    %if "&step1."="Y" %then %do;
    
      data upload_folders(keep=dir_lvl31 summary_path_file s960_upload_path_file);
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
        replacer_filename=catx('_', compress(put(today(),yymmdd10.),,'kd'), compress(put(time(), hhmm.),,'kd'), "summary.xlsx");
        summary_path_file=catx('\', dir_lvl31, replacer_filename);
        s960_upload_filename=cats( 'S960_', region, '_', product_line_group, '_', dir2, '_upload_',  compress(put(today(),yymmdd10.),,'kd'), '_', compress(put(time(), hhmm.),,'kd'), ".txt");
        s960_upload_path_file=catx('\', dir_lvl31, s960_upload_filename);
      run;

    %end;

    %if "&step1."="Y" %then %do;  
      %s960_create_upload_mat_perc(forecast_report_file=%quote(&forecast_report_file.),
                                            region=&region.,
                                            mat_div=&mat_div.,
                                            season=&season.,
                                            product_line_group=&product_line_group.,
                                            seasonality=&seasonality.);

      %s960_generate_upload(forecast_report_file=%quote(&forecast_report_file.),
                                            region=&region.,
                                            mat_div=&mat_div.,
                                            season=&season.,
                                            product_line_group=&product_line_group.,
                                            seasonality=&seasonality.);
    %end;
  %end;

%mend upload_s960;

%upload_s960();