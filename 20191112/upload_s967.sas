/***********************************************************************/
/*Type: Report*/
/*Use: Fill in parameters in metadata.xlsx, sheet=s967 and press run*/
/*Purpose: Create forecast report with 2 seperate steps*/
/*OUT: Replacer list excel report and ZDEMAND format upload file in upload_report_folder (check configuration.sas)*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\cleanup_xlsx_bak_folder.sas";
%include "&sas_applications_folder.\read_metadata.sas";

%macro s967_upload_preparation(forecast_report_file=,
																						region=,
																						Mat_div=,
																						season=,
																						Historical_season=,
																						product_line_group=
																						);

	PROC IMPORT OUT=forecast_report_file 
	            DATAFILE="&forecast_report_file."
	            DBMS=  EXCELCS  REPLACE;
	            SHEET="Variety level fcst"; 
	RUN;

	proc contents data=forecast_report_file out=forecast_contents noprint;
	run;

	data forecast_cols;
		length varnum 8. columnname $32.;
		varnum=4; /*Excel column D*/
		columnname="Country"; 
		output;
		varnum=8; /*Excel column H*/
		columnname="Variety";
		output;
		varnum=31; /*Excel column AE*/
		columnname="Netproposal0";
		output;
		varnum=34; /*Excel column AH*/
		columnname="Netproposal1";
		output;
		varnum=38; /*Excel column AL*/
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

		data orders1(keep=material
										variety
										mat_div
										region
										sub_unit
										actual_sales
										order_year
										order_season
										order_week
										dm_process_stage
										series
										);
		set dmproc.orders_all;
		if mat_div in ("&Mat_div.") and region="&region." and order_season=&historical_season. then output;
	run;

	proc sql;
		create table orders2 as
		select o.* from orders1 o
		right join forecast_report f on o.variety=f.variety and o.region=f.country;
	quit;

	proc sql;
		create table orders3 as
		select * from orders2
		where ^missing(sub_unit) and ^missing(order_year);
	quit;

	proc sql;
		create table orders_average_series as
		select region, order_season, series, sum(actual_sales) as sum_series from orders3
			group by region, order_season, series;
	quit;

	proc sql;
		create table orders_average_series_week as
		select region, order_season, series, order_year, order_week, sum(actual_sales) as sum_series_week 
		from orders3 
		group by region, order_season, series, order_year, order_week;
	quit;

	proc sql;
		create table series_per_week as
		select a.region, a.order_season, a.series, a.order_year, a.order_week, 
			a.sum_series_week, b.sum_series, put(order_year, 4.)||put(order_week, z2.) as yearweek, 
			coalesce((a.sum_series_week/b.sum_series),0) as series_week_percentage
			from orders_average_series_week a
		left join orders_average_series b on a.region=b.region and a.order_season=b.order_season and a.series=b.series;
	quit;

	proc transpose data=series_per_week out=series_per_week1(drop=_name_) prefix=W;
		by region series;
		id yearweek;
		var series_week_percentage;
	run;

	proc sql;
		create table orders_aggr_var as
		select region, order_season, variety, sum(actual_sales) as sum_variety
		from orders3 
		group by region, order_season, variety;
	quit;

	proc sql;
		create table orders_aggr_week as
		select region, order_season, variety, order_year, order_week, sum(actual_sales) as sum_week 
		from orders3 
		group by region, order_season, variety, order_year, order_week;
	quit;

	proc sql;
		create table orders_per_week as
		select w.region, w.order_season, w.variety, w.order_year, w.order_week, w.sum_week, s.sum_variety, w.sum_week/s.sum_variety as week_percentage 
		from orders_aggr_week w
		left join orders_aggr_var s on w.region=s.region and w.variety=s.variety and w.order_season=s.order_season
		order by w.region, w.order_season, w.variety, w.order_year, w.order_week;
	quit;

	%let season_start_week=%scan(&seasonality.,1,'-');
	data all_weeks (drop=i);
	length order_year order_week 8.;
		do i= 1 to 52;
			if i=>&season_start_week. then order_year=&Historical_season.;
				else order_year=&Historical_season.+1;
			order_week=i;
			output;
		end;
	run;

	proc sql;
		create table all_week_combination as
		select * from all_weeks a
		join (select distinct country as region, variety, total_demand from forecast_report) on 1=1
		order by region, variety, order_year, order_week;
	quit;

%mend s967_upload_preparation;

%macro s967_create_upload_mat_perc(forecast_report_file=,
																						region=,
																						Mat_div=,
																						season=,
																						Historical_season=,
																						product_line_group=
																						);

	%s967_upload_preparation(forecast_report_file=%quote(&forecast_report_file.),
																							region=&region.,
																							Mat_div=&Mat_div.,
																							season=&season.,
																							Historical_season=&Historical_season.,
																							product_line_group=&product_line_group.);

	proc sql;
		create table orders_per_week1 as
		select a.region, a.variety, a.total_demand, a.order_year, a.order_week, 
						coalesce(b.order_season, &Historical_season.) as order_season, 
						coalesce(b.sum_week, 0) as sum_week,
						(select distinct c.sum_variety from orders_per_week c where a.variety=c.variety and ^missing(c.sum_variety)) as sum_variety,
						coalesce(b.week_percentage, 0) as week_percentage,
						put(a.order_week, z2.) as yearweek 
		from all_week_combination a
		left join orders_per_week b on a.region=b.region and a.variety=b.variety and a.order_year=b.order_year and a.order_week=b.order_week;
	quit;



	proc transpose data=orders_per_week1 out=orders_per_week2(drop=_name_) prefix=W_;
		by region variety total_demand;
		id yearweek;
		var week_percentage;
	run;

	data orders_per_week3;
		length sum 8.;
		set orders_per_week2;
		sum=sum(of w_:);
	run;

	proc sql;
		create table orders_per_week4 as
		select b.species, b.series, c.variety_description as variety_name, b.current_plc as variety_plc,
					case when sum=0 then 'NO_HISTORY' end as REPLACE, a.* from orders_per_week3 a
		left join dmproc.PMD_Assortment b on a.region=b.region and a.variety=b.variety
		left join (select distinct variety, variety_description from dmimport.material_class_table) c on a.variety=c.variety
		order by a.variety;
	quit;

	data series_average;
		set series_per_week1;
		sum=sum(of w:);
	run;

	data series_average1;
		retain region series sum;
		set series_average;
	run;

	proc sql;
		create table replace as
		select a.*, b.sum as average_sum from orders_per_week4 a 
		left join series_average1 b on a.series=b.series;
	quit;

	data replace1;
		retain region	Species	Series	variety	variety_name	variety_plc	REPLACE	average_sum total_demand	sum;
		set replace;
	run;

	proc sql noprint; 
		select replacer_path_file, dir_lvl31 into :replacer_path_file trimmed, :cleanup_folder trimmed from upload_folders;
	quit;

	x "del &replacer_path_file.";

	proc export 
	  data=replace1
	  dbms=xlsx 
	  outfile="&replacer_path_file." replace;
		sheet="replace";
	run;

	proc export 
	  data=series_average1
	  dbms=xlsx 
	  outfile="&replacer_path_file.";
		sheet="series average";
	run;

	%cleanup_xlsx_bak_folder(cleanup_folder=%str(&cleanup_folder.));

%mend s967_create_upload_mat_perc;

%macro s967_generate_upload(forecast_report_file=,
											region=,
											Mat_div=,
											season=,
											Historical_season=,
											product_line_group=,
											Split_configuration_file=
											);

	%s967_upload_preparation(forecast_report_file=%quote(&forecast_report_file.),
												region=&region.,
												Mat_div=&Mat_div.,
												season=&season.,
												Historical_season=&Historical_season.,
												product_line_group=&product_line_group.);

	%let season0=&season.;
	%let season1=%eval(&season.+1);
	%let season2=%eval(&season.+2);

	data var_net_amounts1(drop=netproposal0 netproposal1 netproposal2);
		length season 8.;
		set forecast_report;
		season=&season0.;
		var_proposal=netproposal0;
		output;
		season=&season1.;
		var_proposal=netproposal1;
		output;
		season=&season2.;
		var_proposal=netproposal2;
		output;
	run;

	PROC IMPORT OUT=var_replace_raw
		DATAFILE="&Split_configuration_file."
		DBMS=  XLSX  REPLACE;
		sheet="replace";
	RUN;

	data var_replace(keep=region variety replace_variety) var_noreplace(keep=region variety) var_average(keep=region variety);
		set var_replace_raw(rename=(replace=_replace));
		if ^missing(_replace) then do;
			if strip(_replace)="AVERAGE" then do; 
				output var_average;
			end; else do;
				replace_variety=input(strip(_replace),8.);
				output var_replace;
			end;
		end; else do;
			output var_noreplace;
		end;
	run;

	proc sql;
		create table orders_per_week_norep as
		select a.region, a.variety, b.order_year, b.order_week, b.sum_week, b.sum_variety, b.week_percentage from var_noreplace a
		left join orders_per_week b on a.variety=b.variety;
	quit;

	proc sql;
		create table orders_per_week_rep as
		select a.region, a.variety, b.order_year, b.order_week, b.sum_week, b.sum_variety, b.week_percentage from var_replace a
		left join orders_per_week b on a.replace_variety=b.variety;
	quit;

	proc sql;
		create table var_average1 as 
		select a.region, a.variety, b.series from var_average a
			left join dmproc.PMD_Assortment b on a.region=b.region and a.variety=b.variety;
	quit;

	proc sql;
		create table orders_per_week_average as
		select  a.region, a.variety, b.order_year, b.order_week, b.sum_series_week as sum_week, b.sum_series as sum_variety, b.series_week_percentage as week_percentage from var_average1 a
		left join series_per_week b on a.region=b.region and a.series=b.series;
	quit;

	data orders_per_week1;
		set orders_per_week_norep
				orders_per_week_rep
				orders_per_week_average;
	run;

	data final_distribution;
		set orders_per_week1(rename=(week_percentage=total_percentage));
	run;

	proc sql;
		create table final_distribution1 as
		select a.region, a.variety, a.order_year, a.order_week, coalesce(b.total_percentage,0) as total_percentage
		from all_week_combination a
		left join final_distribution b on a.region=b.region and a.variety=b.variety and a.order_year=b.order_year and a.order_week=b.order_week;
	quit;

	proc sql;
		create table final_distribution2 as
		select a.*, b.netproposal0 as var_proposal0, b.netproposal1 as var_proposal1, b.netproposal2 as var_proposal2, c.variety_desc
		from final_distribution1 a
		left join forecast_report b on a.variety=b.variety
		left join dmimport.variety_class_table c on a.variety=c.variety;
	quit;

	proc sql;
		create table final_distribution3 as
		select a.*, b.current_plc as variety_plc
		from final_distribution2 a
		left join dmproc.pmd_assortment b on a.variety=b.variety and a.region=b.region;
	quit;

	data final_distribution4;
		set final_distribution3(rename=(order_year=_order_year));
		Confirmed_Sales_Forecast=round(var_proposal0*total_percentage,1);
		order_year=_order_year+&year_offset.;
			output;
		Confirmed_Sales_Forecast=round(var_proposal1*total_percentage,1);
		order_year=_order_year+&year_offset.+1;
			output;
		Confirmed_Sales_Forecast=round(var_proposal2*total_percentage,1);
		order_year=_order_year+&year_offset.+2;
			output;
	run;

	data final_distribution5;
		set final_distribution4;
		if variety_plc^='G2' and Confirmed_Sales_Forecast>0 and ^missing(Confirmed_Sales_Forecast) then output;
	run;

	proc sql;
		create table final_distribution6 as
		select 'S967' as Info_Str,
						0 as Month, 
						put(order_year,4.)||put(order_week,z2.) as Week, 
						0 as Period, 
						'ZF' as Div,
 						'NL01' as Sales_Org,
						'NL81' as Sales_Off,
						'NLUC' as Plant,
						a.variety as Variety,
						a.variety as Material,
						variety_desc as Material_Description,
						'+++' as Sal, 
						'++++++++++' as SRep,
						'++++++++++' as Sold_to_Party,
						Confirmed_Sales_Forecast,
						0 as Returns_Planned,
						0 as Market_Uncertainity,
						Confirmed_Sales_Forecast as Confirmed_Sales_Plan,
						'URC' as Base_Unit 
		from final_distribution5 a
		order by a.variety, Week;
	quit;

	proc sql noprint; 
		select s967_upload_path_file into :s967_upload_path_file trimmed from upload_folders;
	quit;

	PROC EXPORT DATA=final_distribution6
   OUTFILE="&s967_upload_path_file."
       DBMS=TAB REPLACE;
   PUTNAMES=YES;
	RUN;

%mend s967_generate_upload;

%macro upload_s967();

	%read_metadata(sheet=s967);

	proc sql noprint;
		select count(*) into :report_cnt from dmimport.s967_md;
	quit;
	
	%do ii=1 %to &report_cnt.;


		data _null_;
			set dmimport.s967_md;
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

			data upload_folders(keep=dir_lvl31 replacer_path_file s967_upload_path_file);
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
				s967_upload_filename=cats( 'S967_', region, '_', product_line_group, '_', dir2, '_upload_',  compress(put(today(),yymmdd10.),,'kd'), '_', compress(put(time(), hhmm.),,'kd'), ".txt");
				s967_upload_path_file=catx('\', dir_lvl31, s967_upload_filename);
			run;

			%let year_offset=%eval(&season.-&historical_season.);

		%end;

		%if "&step1."="Y" %then %do;	
			%s967_create_upload_mat_perc(forecast_report_file=%quote(&forecast_report_file.),
																						region=&region.,
																						Mat_div=&Mat_div.,
																						season=&season.,
																						Historical_season=&Historical_season.,
																						product_line_group=&product_line_group.);
		%end;

		%if "&step2."="Y" %then %do;
			%s967_generate_upload(forecast_report_file=%quote(&forecast_report_file.),
																						region=&region.,
																						Mat_div=&Mat_div.,
																						season=&season.,
																						Historical_season=&Historical_season.,
																						product_line_group=&product_line_group.,
																						Split_configuration_file=%quote(&Split_configuration_file.));
		%end;
	%end;

%mend upload_s967;

%upload_s967();