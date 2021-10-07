/***********************************************************************/
/*Type: Report*/
/*Use: Fill in parameters in metadata.xlsx, sheet=s965_s966 and press run*/
/*Purpose: Create forecast report with 2 seperate steps*/
/*OUT: Replacer list excel report and ZDEMAND format upload file in in upload_report_folder (check configuration.sas)*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\cleanup_xlsx_bak_folder.sas";
%include "&sas_applications_folder.\read_metadata.sas";

%macro s965_s966_upload_preparation(forecast_report_file=,
																						region=,
																						material_division=,
																						season=,
																						historical_season=,
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

	proc sql;
		create table varieties_not_in_pmd as
		select variety from forecast_report where variety not in (select variety from dmproc.pmd_assortment where region='FPS');
	quit;

	proc sql;
		create table varieties_not_in_pmd1 as
		select a.*,b.current_plc as variety_plc, c.total_demand from varieties_not_in_pmd a
		left join dmproc.pmd_assortment b on a.variety=b.variety 
		left join forecast_report c on a.variety=c.variety;
	quit;

	proc sql;
		create table varieties_not_in_fps as
		select variety from forecast_report where variety not in (select variety from dmimport.fps_assortment where region='FPS' and material_division in ('6B', '6C'));
	quit;

	proc sql;
		create table varieties_not_in_fps1 as
		select a.*,b.current_plc as variety_plc, c.total_demand from varieties_not_in_fps a
		left join dmproc.pmd_assortment b on a.variety=b.variety 
		left join forecast_report c on a.variety=c.variety;
	quit;

	proc sql;
		create table varieties_without_material as
		select a.variety, b.variety_name, b.current_plc as variety_plc, a.total_demand, c.material,  c.curr_mat_plc as material_plc 
		from forecast_report a
		left join dmproc.pmd_assortment b on a.country=b.region and a.variety=b.variety
		left join dmimport.fps_assortment c on a.variety=c.variety;
	quit;

	proc sql;
		create table varieties_without_material1 as
		select a.* 
		from varieties_without_material a
		left join (select variety, count(*) as g2_cnt from varieties_without_material where material_plc='G2' group by variety) b on a.variety=b.variety
		left join (select variety, count(*) as not_g2_cnt from varieties_without_material where material_plc^='G2' group by variety) c on a.variety=c.variety
		where missing(a.material) or (g2_cnt>1 and not_g2_cnt=0);
	quit;

	proc sql;
		create table forecast_report0 as
		select * from forecast_report 
		where 
			variety not in (select variety from varieties_not_in_pmd) and 
			variety not in (select variety from varieties_not_in_fps);
	quit;

	proc sql;
		create table forecast_report1 as
		select a.variety, a.country as region, a.total_demand, a.Netproposal0, a.Netproposal1, a.Netproposal2, b.material, b.material_division as mat_div, &historical_season. as order_season, b.dm_process_stage
		from forecast_report0 a
		left join dmimport.fps_assortment b on a.variety=b.variety and material_division in ('6B', '6C');
	quit;

	proc sql;
		create table materials_no_delivery_window as
		select material from forecast_report1 where material not in (select material from dmimport.delivery_window);
	quit;

	proc sql;
		create table materials_no_delivery_window1 as
		select a.*, b.curr_mat_plc as material_plc, b.variety, c.current_plc as variety_plc, d.total_demand from materials_no_delivery_window a
		left join dmimport.fps_assortment b on a.material=b.material
		left join dmproc.pmd_assortment c on b.variety=c.variety 
		left join forecast_report d on b.variety=d.variety;
	quit;

	proc sql;
		create table varieties_no_delivery_window as
		select variety from forecast_report where variety not in (select variety from dmimport.delivery_window);
	quit;

	proc sql;
		create table varieties_no_delivery_window1 as
		select a.*, b.current_plc as variety_plc, c.total_demand from varieties_no_delivery_window a
		left join dmproc.pmd_assortment b on a.variety=b.variety 
		left join forecast_report c on a.variety=c.variety;
	quit;

	data orders1(keep=material
										variety
										mat_div
										region
										sub_unit
										actual_sales
										order_year
										order_season
										order_week
										);
		set dmproc.orders_all;
		if upcase(mat_div) in ("6B", "6C") and region="&region." and order_season=&historical_season. then output;
	run;

	proc sql;
		create table orders2 as
		select f.region, f.variety, f.dm_process_stage, 
					 o.material, o.mat_div, o.sub_unit, o.actual_sales, o.order_year, o.order_season, o.order_week 
		from orders1 o
		right join forecast_report1 f on o.region=f.region and o.variety=f.variety;
	quit;

	proc sql;
		create table orders3 as
		select * from orders2
		where ^missing(sub_unit) and ^missing(order_year);
	quit;

	proc sql;
		create table orders4 as
		select a.*, b.species, b.series, b.variety_name, c.material_basic_description_en as material_name, c.PF_for_sales_text from orders3 a
			left join dmproc.PMD_Assortment b on a.region=b.region and a.variety=b.variety
			left join dmimport.fps_assortment c on a.region=c.region and a.variety=c.variety and a.material=c.material;
	quit;

	proc sql;
		create table orders_average_series as
		select region, order_season, series, PF_for_sales_text, sum(actual_sales) as sum_series from orders4
			group by region, order_season, series, PF_for_sales_text;
	quit;

	proc sql;
		create table orders_average_series_week as
		select region, order_season, series, PF_for_sales_text, order_year, order_week, sum(actual_sales) as sum_series_week 
		from orders4 
		group by region, order_season, series, PF_for_sales_text, order_year, order_week;
	quit;

	proc sql;
		create table series_per_week as
		select a.region, a.order_season, a.series, a.PF_for_sales_text, a.order_year, a.order_week, 
			a.sum_series_week, b.sum_series, put(order_year, 4.)||put(order_week, z2.) as yearweek, 
			coalesce((a.sum_series_week/b.sum_series),0) as series_week_percentage
			from orders_average_series_week a
		left join orders_average_series b on a.region=b.region and a.order_season=b.order_season and a.series=b.series and a.PF_for_sales_text=b.PF_for_sales_text;
	quit;

	proc transpose data=series_per_week out=series_per_week1(drop=_name_) prefix=W;
		by region series PF_for_sales_text;
		id yearweek;
		var series_week_percentage;
	run;

	proc sql;
		create table orders_aggr_season as
		select region, order_season, variety, sum(actual_sales) as sum_variety
		from orders3 
		group by region, order_season, variety;
	quit;

	proc sql;
		create table orders_aggr_season1 as
		select a.region, a.order_season, a.variety, a.total_demand, coalesce(sum_variety,0) as sum_variety 
		from (select distinct region, order_season, variety, total_demand from forecast_report1) a
		left join orders_aggr_season b on a.region=b.region and a.order_season=b.order_season and a.variety=b.variety;
	quit;

	proc sql;
		create table orders_aggr_mat as
		select region, order_season, variety, mat_div, material, sum(actual_sales) as sum_material 
		from orders3 
		group by region, order_season, variety, mat_div, material;
	quit;

	proc sql;
		create table orders_aggr_mat1 as
		select a.region, a.order_season, a.variety, a.total_demand, a.mat_div, a.material, coalesce(sum_material,0) as sum_material 
		from  forecast_report1 a
		left join orders_aggr_mat b on a.region=b.region and a.order_season=b.order_season and a.variety=b.variety and a.material=b.material and a.mat_div=b.mat_div;
	quit;

	proc sql;
		create table orders_per_mat as
		select w.region, w.order_season, w.variety, w.mat_div, w.material, w.sum_material, w.total_demand, s.sum_variety, 
					w.sum_material/s.sum_variety as material_percentage, . as adjusted_percentage  
		from orders_aggr_mat1 w
		left join orders_aggr_season1 s on w.region=s.region and w.variety=s.variety and w.order_season=s.order_season
		order by w.region, w.order_season, w.variety, w.mat_div, w.material;
	quit;

	proc sql;
		create table orders_per_mat1 as
		select a.*, b.species, b.series, b.variety_name, b.current_plc as variety_plc, c.material_basic_description_en as material_name, c.PF_for_sales_text, c.curr_mat_plc as material_plc
		from orders_per_mat a
		left join dmproc.PMD_Assortment b on a.region=b.region and a.variety=b.variety
		left join dmimport.fps_assortment c on a.region=c.region and a.variety=c.variety and a.material=c.material
		order by a.variety, a.material;
	quit;

	proc sql;
		create table orders_aggr_week as
		select region, order_season, variety, mat_div, material, order_year, order_week, sum(actual_sales) as sum_week 
		from orders3 
		group by region, order_season, variety, mat_div, material, order_year, order_week;
	quit;

	proc sql;
		create table orders_per_week as
		select w.region, w.order_season, w.variety, w.mat_div, w.material, w.order_year, w.order_week, w.sum_week, s.sum_material, w.sum_week/s.sum_material as week_percentage 
		from orders_aggr_week w
		left join orders_aggr_mat s on w.region=s.region and w.variety=s.variety and w.order_season=s.order_season and w.material=s.material
		order by w.region, w.order_season, w.variety, w.mat_div, w.material, w.order_year, w.order_week;
	quit;

	%let season_start_week=%scan(&seasonality.,1,'-');
	data all_weeks (drop=i);
	length order_year order_week 8.;
		do i= 1 to 52;
			if i=>&season_start_week. then order_year=&historical_season.;
				else order_year=&historical_season.+1;
			order_week=i;
			output;
		end;
	run;

	proc sql;
		create table all_week_combination as
		select * from all_weeks a
		join (select distinct region, variety, mat_div, material, total_demand from forecast_report1) on 1=1
		order by region, variety, mat_div, material, order_year, order_week;
	quit;

%mend s965_s966_upload_preparation;

%macro s965_s966_create_upload_mat_perc(forecast_report_file=,
																						region=,
																						material_division=,
																						season=,
																						historical_season=,
																						product_line_group=
																						);

	%s965_s966_upload_preparation(forecast_report_file=%quote(&forecast_report_file.),
																						region=&region.,
																						material_division=%quote(&material_division.),
																						season=&season.,
																						historical_season=&historical_season.,
																						product_line_group=&product_line_group.);

	proc sql;
		create table orders_per_week1 as
		select a.region, a.variety, a.mat_div, a.material, a.order_year, a.order_week, a.total_demand, 
						coalesce(b.order_season, &historical_season.) as order_season, 
						coalesce(b.sum_week, 0) as sum_week,
						(select distinct c.sum_material from orders_per_week c where a.material=c.material and ^missing(c.sum_material)) as sum_material,
						coalesce(b.week_percentage, 0) as week_percentage,
						put(a.order_week, z2.) as yearweek 
		from all_week_combination a
		left join orders_per_week b on a.region=b.region and a.variety=b.variety and a.mat_div=b.mat_div and a.material=b.material and a.order_year=b.order_year and a.order_week=b.order_week;
	quit;

	proc transpose data=orders_per_week1 out=orders_per_week2(drop=_name_) prefix=W_;
		by region variety mat_div material total_demand;
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
		select b.species, b.series, b.variety_name, b.current_plc as variety_plc, 
					c.material_basic_description_en as material_name, c.PF_for_sales_text, c.curr_mat_plc as material_plc, c.replaced_by as replaced_by_suggestion,
					case when sum=0 then 'NO_HISTORY' end as REPLACE, a.* from orders_per_week3 a
		left join dmproc.PMD_Assortment b on a.region=b.region and a.variety=b.variety
		left join dmimport.fps_assortment c on a.region=c.region and a.variety=c.variety and a.material=c.material
		order by a.variety, a.mat_div, a.material;
	quit;

	proc sql noprint; 
		select replacer_path_file, dir_lvl31 into :replacer_path_file trimmed, :cleanup_folder trimmed from upload_folders;
	quit;

	x "del &replacer_path_file.";

	proc sql;
		create table adjust as
		select a.*, b.var_count from orders_per_mat1 a
		left join (select variety, count(*) as var_count from orders_per_mat1 group by variety) b on a.variety=b.variety;
	quit;

	quit;
	data adjust1;
		retain region	order_season	variety	var_count	variety_plc	material	mat_div	material_plc	PF_for_sales_text	 sum_material 	 sum_variety 	 total_demand 	 material_percentage 	adjusted_percentage	Species	Series	Variety_name	material_name;
		set adjust;
	run;

	proc export 
	  data=adjust1 
	  dbms=xlsx 
	  outfile="&replacer_path_file." replace;
		sheet="adjust";
	run;

	data averages_info;
		set series_per_week1;
		sum=sum(of W:);
	run;

	data averages_info1;
		retain region series pf_for_sales_text sum;
		set averages_info;
	run;

	proc sql;
		create table replace as
		select a.*, b.sum as average_sum from orders_per_week4 a
		left join averages_info1 b on a.series=b.series and a.pf_for_sales_text=b.pf_for_sales_text;
	quit;

	data replace1;
		retain region	Species	Series	variety	Variety_name	variety_plc	material	material_name	mat_div	PF_for_sales_text	material_plc	replaced_by_suggestion	REPLACE	average_sum total_demand	sum;
		set replace;
	run;

	proc export 
	  data=replace1
	  dbms=xlsx 
	  outfile="&replacer_path_file.";
		sheet="replace";
	run;

	proc export 
	  data=averages_info1
	  dbms=xlsx 
	  outfile="&replacer_path_file.";
		sheet="averages_info";
	run;

	proc export 
	  data=varieties_not_in_pmd1
	  dbms=xlsx 
	  outfile="&replacer_path_file.";
		sheet="varieties_not_in_pmd";
	run;

	proc export 
	  data=varieties_not_in_fps1
	  dbms=xlsx 
	  outfile="&replacer_path_file.";
		sheet="varieties_not_in_fps";
	run;

	proc export 
	  data=varieties_no_delivery_window1
	  dbms=xlsx 
	  outfile="&replacer_path_file.";
		sheet="varieties_no_delivery_window";
	run;

	proc export 
	  data=materials_no_delivery_window1
	  dbms=xlsx 
	  outfile="&replacer_path_file.";
		sheet="materials_no_delivery_window";
	run;

	proc export 
	  data=varieties_without_material1
	  dbms=xlsx 
	  outfile="&replacer_path_file.";
		sheet="varieties_without_material";
	run;

	%cleanup_xlsx_bak_folder(cleanup_folder=%str(&cleanup_folder.));

%mend s965_s966_create_upload_mat_perc;

%macro s965_s966_generate_upload(forecast_report_file=,
											region=,
											material_division=,
											season=,
											historical_season=,
											product_line_group=,
											split_configuration_file=
											);

	%let season0=&season.;
	%let season1=%eval(&season.+1);
	%let season2=%eval(&season.+2);

	%s965_s966_upload_preparation(forecast_report_file=%quote(&forecast_report_file.),
																						region=&region.,
																						material_division=%quote(&material_division.),
																						season=&season.,
																						historical_season=&historical_season.,
																						product_line_group=&product_line_group.);

	PROC IMPORT OUT=mat_percentage_adjusted
		DATAFILE="&split_configuration_file."
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
		select d.*, p.mat_dist_percentage from mat_percentage_adjusted1 d 
		left join (select e.*, (e.material_percentage/s.sum_var) as mat_dist_percentage from mat_percentage_adjusted1 e 
		left join (select variety, sum(material_percentage) as sum_var from mat_percentage_adjusted1 where missing(adjusted_percentage) group by variety) s on e.variety=s.variety
		where missing(e.adjusted_percentage)) p on d.material=p.material;
	quit;

	proc sql;
		create table mat_adjusted_to_dist as 
		select variety, sum(material_percentage-adjusted_percentage) as perc_diff from mat_percentage_adjusted1 where ^missing(adjusted_percentage) group by variety;
	quit;

	proc sql;
		create table mat_perc_distributed as
		select region, variety, material, mat_div, coalesce(distributed_mat_perc, adjusted_percentage, material_percentage) as final_mat_percentage from 
		(select a.*, b.perc_diff, a.material_percentage+(b.perc_diff*a.mat_dist_percentage) as distributed_mat_perc from mat_percentage_dist1 a
		left join mat_adjusted_to_dist b on a.variety=b.variety and missing(a.adjusted_percentage));
	quit;

	proc sql;
		create table mat_net_amounts as
		select a.*, (b.netproposal0*a.final_mat_percentage) as mat_proposal0,
								(b.netproposal1*a.final_mat_percentage) as mat_proposal1,
								(b.netproposal2*a.final_mat_percentage) as mat_proposal2
		from mat_perc_distributed a 
		left join forecast_report0 b on a.variety=b.variety;
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
		DATAFILE="&split_configuration_file."
		DBMS=  XLSX  REPLACE;
		sheet="replace";
	RUN;

	data mat_replace(keep=region variety material mat_div replace_material) mat_average(keep=region variety mat_div material) mat_noreplace(keep=region variety mat_div material);
		set mat_replace_raw(rename=(replace=_replace));
		if ^missing(_replace) then do;
			if strip(_replace)="AVERAGE" then do; 
				output mat_average;
			end; else do;
				replace_material=input(strip(_replace),8.);
				output mat_replace;
			end;
		end; else do;
			output mat_noreplace;
		end;
	run;

	proc sql;
		create table orders_per_week_norep as
		select b.* from mat_noreplace a
		left join orders_per_week b on a.material=b.material;
	quit;

	proc sql;
		create table orders_per_week_rep as
		select a.region, a.variety, a.mat_div, a.material, b.order_year, b.order_week, b.sum_week, b.sum_material, b.week_percentage from mat_replace a
		left join orders_per_week b on a.replace_material=b.material;
	quit;

	proc sql;
		create table mat_average1 as 
		select a.variety, a.mat_div, a.material, b.series, c.PF_for_sales_text from mat_average a
			left join dmproc.PMD_Assortment b on a.region=b.region and a.variety=b.variety
			left join dmimport.fps_assortment c on a.region=c.region and a.variety=c.variety and a.material=c.material;
	quit;

	proc sql;
		create table orders_per_week_average as
		select a.variety, a.mat_div, a.material, b.* from mat_average1 a
		left join series_per_week b on a.series=b.series and a.PF_for_sales_text=b.PF_for_sales_text;
	quit;

	data orders_per_week1;
		set orders_per_week_norep
				orders_per_week_rep
				orders_per_week_average(drop=series PF_for_sales_text rename=(sum_series_week=sum_week sum_series=sum_material series_week_percentage=week_percentage));
	run;

	data orders_per_week2(drop=rc);
		set orders_per_week1;
		length status 8.;
		if _n_=1 then do;
			declare hash delivery_window(dataset: 'dmimport.delivery_window');
				rc=delivery_window.DefineKey ('material', 'delivery_week_year', 'delivery_week_week');
				rc=delivery_window.DefineData ('status');
				rc=delivery_window.DefineDone();
		end;
		delivery_week_week=order_week;
		delivery_week_year=order_year+&year_offset.;/*take delivery schedule from current season*/
		if mat_div='6B' then do;
			status=0;/*should be closed by default*/
			rc=delivery_window.find();
		end;
		if mat_div='6C' then status=1; /*always open for 6C*/
	run;

	data orders_per_week_closed orders_per_week_open;
		set orders_per_week2;
		if status=1 then output orders_per_week_open;
		if status=0 then output orders_per_week_closed;
	run;

	proc sql;
		create table orders_per_week_closed_aggr as
		select region, variety, material, order_season, sum(week_percentage) as sum_of_closed_weeks from orders_per_week_closed group by region, variety, material, order_season;
	quit;

	proc sql;
		create table orders_per_week_open_weeks as
		select region, variety, material, order_season, count(*) as count_open_weeks from orders_per_week_open where week_percentage^=0 group by region, variety, material, order_season;
	quit;

	proc sql;
		create table orders_per_week_closed_aggr1 as
		select a.*, a.sum_of_closed_weeks/o.count_open_weeks as redistributed_closed_weeks from orders_per_week_closed_aggr a
		left join orders_per_week_open_weeks o on a.region=o.region and a.variety=o.variety and a.material=o.material and a.order_season=o.order_season;
	quit;

	proc sql;
		create table final_distribution as
		select 
			o.*,
			case 
				when coalesce(o.week_percentage,0)=0 then 0
				else a.redistributed_closed_weeks 
			end as redistributed_closed_weeks, 
			case 
				when coalesce(o.week_percentage,0)=0 then 0
				else coalesce(o.week_percentage,0)+coalesce(a.redistributed_closed_weeks,0)
			end as total_percentage
		from orders_per_week_open o
		left join orders_per_week_closed_aggr1 a on a.region=o.region and a.variety=o.variety and a.material=o.material and a.order_season=o.order_season;
	quit;

	proc sql;
		create table final_distribution1 as
		select a.region, a.variety, a.mat_div, a.material, a.order_year, a.order_week, coalesce(b.total_percentage,0) as total_percentage
		from all_week_combination a
		left join final_distribution b on a.region=b.region and a.variety=b.variety and a.material=b.material and a.order_year=b.order_year and a.order_week=b.order_week;
	quit;

	proc sql;
		create table final_distribution2 as
		select a.*, b.final_mat_percentage, b.mat_proposal0, b.mat_proposal1, b.mat_proposal2, c.material_basic_description_en, c.sub_unit, d.current_plc as variety_plc 
		from final_distribution1 a
		left join mat_net_amounts b on a.region=b.region and a.variety=b.variety and a.material=b.material
		left join dmimport.fps_assortment c on a.region=c.region and a.variety=c.variety and a.material=c.material
		left join dmproc.pmd_assortment d on a.region=d.region and a.variety=d.variety;
	quit;
	
	data final_distribution3;
		set final_distribution2(rename=(order_year=_order_year));
		Confirmed_Sales_Forecast=round(mat_proposal0*total_percentage,1);
		Confirmed_Sales_Forecast_box=round(mat_proposal0*total_percentage/sub_unit,1);
		order_year=_order_year+&year_offset.;
			output;
		Confirmed_Sales_Forecast=round(mat_proposal1*total_percentage,1);
		Confirmed_Sales_Forecast_box=round(mat_proposal1*total_percentage/sub_unit,1);
		order_year=_order_year+&year_offset.+1;
			output;
		Confirmed_Sales_Forecast=round(mat_proposal2*total_percentage,1);
		Confirmed_Sales_Forecast_box=round(mat_proposal2*total_percentage/sub_unit,1);
		order_year=_order_year+&year_offset.+2;
			output;
	run;

	data final_distribution4;
		set final_distribution3;
		if variety_plc^='G2' then output;
	run;

	data final_distribution4_6B final_distribution4_6C;
		set final_distribution4;
		if mat_div='6B' and Confirmed_Sales_Forecast_box>0 then output final_distribution4_6B;
		if mat_div='6C' and Confirmed_Sales_Forecast>0 then output final_distribution4_6C;
	run;

	proc sql;
		create table final_distribution5_6B as
		select 'S966' as Info_Str,
						0 as Month, 
						put(order_year,4.)||put(order_week,z2.) as Week, 
						0 as Period, 
						'ZF' as Div,
 						'NL02' as Sales_Org,
						'NL50' as Sales_Off,
						'NLYP' as Plant,
						a.variety as Variety,
						a.material as Material,
						material_basic_description_en as Material_Description,
						'+++' as Sales_Grp, 
						'++++++++++' as SRep,
						'++++++++++' as Sold_to_Party,
						Confirmed_Sales_Forecast_box as Confirmed_Sales_Forecast,
						0 as Returns_Planned,
						0 as Market_Uncertainity,
						Confirmed_Sales_Forecast_box as Confirmed_Sales_Plan,
						'BOX' as Base_Unit 
		from final_distribution4_6B a
		order by a.variety, a.material, Week;
	quit;

	proc sql;
		create table final_distribution5_6C as
		select 'S965' as Info_Str,
						0 as Month, 
						put(order_year,4.)||put(order_week,z2.) as Week, 
						0 as Period, 
						'ZF' as Div,
 						'NL02' as Sales_Org,
						'NL50' as Sales_Off,
						'NLUC' as Plant,
						a.variety as Variety,
						a.material as Material,
						material_basic_description_en as Material_Description,
						'+++' as Sales_Grp, 
						'++++++++++' as SRep,
						'++++++++++' as Sold_to_Party,
						Confirmed_Sales_Forecast,
						0 as Returns_Planned,
						0 as Market_Uncertainity,
						Confirmed_Sales_Forecast as Confirmed_Sales_Plan,
						'URC' as Base_Unit 
		from final_distribution4_6C a 
		order by a.variety, a.material, Week;
	quit;

	proc sql noprint; 
		select s965_upload_path_file into :s965_upload_path_file trimmed from upload_folders;
		select s966_upload_path_file into :s966_upload_path_file trimmed from upload_folders;
	quit;

	PROC EXPORT DATA=final_distribution5_6B
   OUTFILE="&s966_upload_path_file."
       DBMS=TAB REPLACE;
   PUTNAMES=YES;
	RUN;

		PROC EXPORT DATA=final_distribution5_6C
   OUTFILE="&s965_upload_path_file."
       DBMS=TAB REPLACE;
   PUTNAMES=YES;
	RUN;

%mend s965_s966_generate_upload;


%macro upload_s965_s966();

	%read_metadata(sheet=S965_S966);

	proc sql noprint;
		select count(*) into :report_cnt from dmimport.S965_S966_md;
	quit;
	
	%do ii=1 %to &report_cnt.;

		data _null_;
			set dmimport.S965_S966_md;
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

			data upload_folders(keep=dir_lvl31 replacer_path_file s965_upload_path_file s966_upload_path_file);
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
				s965_upload_filename=cats( 'S965_', region, '_', product_line_group, '_', dir2, '_upload_',  compress(put(today(),yymmdd10.),,'kd'), '_', compress(put(time(), hhmm.),,'kd'), ".txt");
				s965_upload_path_file=catx('\', dir_lvl31, s965_upload_filename);
				s966_upload_filename=cats( 'S966_', region, '_', product_line_group, '_', dir2, '_upload_',  compress(put(today(),yymmdd10.),,'kd'), '_', compress(put(time(), hhmm.),,'kd'), ".txt");
				s966_upload_path_file=catx('\', dir_lvl31, s966_upload_filename);
			run;

			%let year_offset=%eval(&season.-&historical_season.);

		%end;

		%if "&step1."="Y" %then %do;	
			%s965_s966_create_upload_mat_perc(forecast_report_file=%quote(&forecast_report_file.),
																									region=&region.,
																									material_division=%quote(&mat_div.),
																									season=&season.,
																									historical_season=&historical_season.,
																									product_line_group=&product_line_group.);
		%end;

		%if "&step2."="Y" %then %do;
			%s965_s966_generate_upload(forecast_report_file=%quote(&forecast_report_file.),
																									region=&region.,
																									material_division=%quote(&mat_div.),
																									season=&season.,
																									historical_season=&historical_season.,
																									product_line_group=&product_line_group.,
																									split_configuration_file=%str(&Split_configuration_file.));
		%end;
	%end;

%mend upload_s965_s966;

%upload_s965_s966();