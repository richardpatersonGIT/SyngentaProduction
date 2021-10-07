/***********************************************************************/
/*Type: Report*/
/*Use: Fill in parameters in exrapolation_report.xlsx and press run*/
/*Purpose: Create Orders report from Running Sales*/
/*OUT: Excel file to extrapolation_report_folder (check configuration.sas)*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\cleanup_xlsx_bak_folder.sas";

%macro read_extrapolation_metadata();

	PROC IMPORT OUT=extrapolation_report_md_raw
	            DATAFILE="&ex_metadata_file."
	            DBMS=  EXCELCS  REPLACE; 
	RUN;

	data extrapolation_report_md(drop=_:);
		set extrapolation_report_md_raw(rename=(week=_week));
		if ^missing(coalesceC(of _character_)) or ^missing(coalesce(of _numeric_)) then output;
		week=input(_week, best.);
	run;

%mend read_extrapolation_metadata;

%macro extrapolation_extraction(extrapolation_season=, 
																region=, 
																mat_div=,
																mat_div_quote=,
																product_line_group=);


	data extrapolation_report_md1(drop=ii hist_season);
		set extrapolation_report_md(obs=1);
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

		data sales_history_tmp1(drop=sls_org sls_off shipto_cntry order_type rsn_rej_cd _cnf_qty);
			set shw.&dsname. (keep=sls_org sls_off shipto_cntry material variety SchedLine_Cnf_deldte cnf_qty order_type mat_div rsn_rej_cd rename=(cnf_qty=_cnf_qty));
			length unique_code $10. year 8. week 8.;
			unique_code=cats(sls_org, sls_off, shipto_cntry);
			year=&year.;
			week=&week.;
			cnf_qty=input(_cnf_qty, comma15.2);
			if mat_div in (&mat_div_quote.) and ((mat_div in ('6B', '6C') and order_type in ('ZYPD', 'ZFD1', 'ZYPL')) or (mat_div='6A' and order_type='YQOR')) and missing(rsn_rej_cd) and ^missing(SchedLine_Cnf_deldte) then output;
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
				macro_region="&region.";
		end;

		rc=cl.find(); /*gets territory and country from country_lookup*/
		if region=macro_region and mat_div in (&mat_div_quote.) then output;
	run;

	data sales_history2(drop=rc);
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
		If region = 'BI' then rc=bi_assortment.find(); /*gets sub_unit from material_class_table*/

		if ^missing(sub_unit) then do;
			historical_sales=cnf_qty * sub_unit;
		end;

	run;

	data sales_history3(keep=region year week hash_species_name order_season country historical_sales);
		set sales_history2;
		length	season_week_start season_week_end 
					 	Order_season_start order_year order_season order_week Order_yweek order_month macro_extrapolation_season 8.;
		length hash_species_name $29. product_line_group $20. macro_product_line_group $20.;
		retain macro_extrapolation_season macro_product_line_group;
		if _n_=1 then do;
			declare hash pmd_assortment(dataset: 'dmproc.PMD_assortment');
				rc=pmd_assortment.DefineKey ('region', 'variety');
				rc=pmd_assortment.DefineData ('season_week_start', 'season_week_end', 'hash_species_name', 'product_line_group');
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

	proc sql;
		create table sales_history_aggr as
		select region, year, week, hash_species_name, order_season, country, round(sum(historical_sales),1) as country_sales_aggr 
			from sales_history3 
			group by region, year, week, hash_species_name, order_season, country;
	quit;

	proc sql;
		create table country_mid_season_aggr as
		select b.* from extrapolation_report_wks a
		left join sales_history_aggr b on a.mid_season_year=b.year and a.mid_season_week=b.week;
	quit;

	proc sql;
		create table species_mid_season_aggr as
		select region, year, week, hash_species_name, order_season, round(sum(country_sales_aggr),1) as species_sales_aggr from country_mid_season_aggr a
		group by region, year, week, hash_species_name, order_season;
	quit;

	proc sql;
		create table country_end_season_aggr as
		select b.* from extrapolation_report_wk_end a
		left join sales_history_aggr b on a.end_season_year=b.year and a.end_season_week=b.week;
	quit;

	proc sql;
		create table species_end_season_aggr as
		select region, year, week, hash_species_name, order_season, round(sum(country_sales_aggr),1) as species_sales_aggr from country_end_season_aggr a
		group by region, year, week, hash_species_name, order_season;
	quit;

	proc sql;
		create table all_country_weeks as
		select b.*, a.* from 
			(select distinct mid_season_year, mid_season_week from extrapolation_report_md1) a
			inner join (select distinct hash_species_name, country from country_end_season_aggr) b on 1=1;
	quit;

	proc sql;
		create table all_species_weeks as
		select b.*, a.* from  
			(select distinct mid_season_year, mid_season_week from extrapolation_report_md1) a
			inner join (select distinct hash_species_name from species_end_season_aggr) b on 1=1;
	quit;

	proc sql;
	create table country_sales_percentage as
		select 	e.region, 
						"&product_line_group." as product_line_group, 
						propcase(a.hash_species_name) as species, 
						"&mat_div." as mat_div,
						e.order_season as hist_season,
						a.country, 
						a.mid_season_year as mid_year,
						a.mid_season_week as mid_week,
						m.country_sales_aggr as sales, 
						e.country_sales_aggr as end_season_sales,  
						m.country_sales_aggr/e.country_sales_aggr as extrapolation_rate 
		from all_country_weeks a
		left join country_end_season_aggr e on a.hash_species_name=e.hash_species_name and a.country=e.country  
		left join country_mid_season_aggr m on m.hash_species_name=a.hash_species_name and m.country=a.country and m.year=a.mid_season_year and m.week=a.mid_season_week;
	quit;

	proc sql;
	create table species_sales_percentage as
		select 	e.region, 
						"&product_line_group." as product_line_group, 
						propcase(a.hash_species_name) as species, 
						"&mat_div." as mat_div,
						e.order_season as hist_season, 
						a.mid_season_year as mid_year,
						a.mid_season_week as mid_week,
						m.species_sales_aggr as sales, 
						e.species_sales_aggr as end_season_sales,  
						m.species_sales_aggr/e.species_sales_aggr as extrapolation_rate 
		from all_species_weeks a
		left join species_end_season_aggr e on a.hash_species_name=e.hash_species_name  
		left join species_mid_season_aggr m on m.hash_species_name=a.hash_species_name and m.year=a.mid_season_year and m.week=a.mid_season_week;
	quit;

%mend extrapolation_extraction;

%macro extrapolation_report();

	%let extrapolation_start_time=EXTRAPOLATION STARTED: %sysfunc(date(),worddate.). %sysfunc(time(),timeampm.);

	%read_extrapolation_metadata();

	proc sql noprint;
		select hist_season into :extrapolation_season trimmed from extrapolation_report_md where ^missing(hist_season);
		select region into :region trimmed from extrapolation_report_md where ^missing(region);
		select '"'||mat_div||'"' into :mat_div_quote separated by ', ' from extrapolation_report_md where ^missing(mat_div);
		select mat_div into :mat_div separated by ',' from extrapolation_report_md where ^missing(mat_div);
		select product_line_group into :product_line_group trimmed from extrapolation_report_md where ^missing(product_line_group);
	quit;

	%if not %symexist(species) %then %let species=empty;

	%extrapolation_extraction(extrapolation_season=&extrapolation_season., 
															region=&region., 
															mat_div=%quote(&mat_div.),
															mat_div_quote=%quote(&mat_div_quote.),
															product_line_group=&product_line_group.);

	proc sql noprint;
			select mat_div into :mat_div_name separated by '_' from extrapolation_report_md where ^missing(mat_div);
			select catx('_', region, product_line_group, "&mat_div_name.", seasonality, put(hist_season, 4.), cats("wk",put(coalesce(week, 0), z2.))) into :ext_report_name trimmed from extrapolation_report_md where ^missing(region);
	quit;

	data _null_;
		extrapolation_report_file=catx('_', compress(put(today(),yymmdd10.),,'kd'), compress(put(time(), time8.),,'kd'));
		call symput('extrapolation_report_file', strip(extrapolation_report_file));
	run;

	%let extrapolation_name=&extrapolation_report_folder.\Extrapolation_&ext_report_name._&extrapolation_report_file..xlsx;

	x "del &extrapolation_name."; 

	proc export 
	  data=species_sales_percentage
	  dbms=xlsx 
	  outfile="&extrapolation_name." replace;
		sheet="Species";
	run;

	proc export 
	  data=country_sales_percentage
	  dbms=xlsx 
	  outfile="&extrapolation_name.";
		sheet="All_countries";
	run;

	proc sql noprint;
		create table countries as
		select distinct country from all_country_weeks;
		select country into :countries separated by '#' from countries;
		select count(*) into: countries_cnt trimmed from countries;
	quit;

	%do ci=1 %to &countries_cnt.;
		%let country=%scan(&countries., &ci., '#');
		data country_report;
			set country_sales_percentage;
			if country="&country." then output;
		run;

		proc export 
		  data=country_report
		  dbms=xlsx 
		  outfile="&extrapolation_name.";
			sheet="&country.";
		run;
	%end;

	proc export 
	  data=extrapolation_report_md
	  dbms=xlsx 
	  outfile="&extrapolation_name.";
		sheet="Variant";
	run;

	%cleanup_xlsx_bak_folder(cleanup_folder=%str(&extrapolation_report_folder.\));

	%let extrapolation_end_time=EXTRAPOLATION ENDED: %sysfunc(date(),worddate.). %sysfunc(time(),timeampm.);

	%put &=extrapolation_report_folder.;
	%put &=extrapolation_end_time.;

%mend extrapolation_report;

%extrapolation_report();



