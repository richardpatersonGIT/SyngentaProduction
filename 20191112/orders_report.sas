/***********************************************************************/
/*Type: Report*/
/*Use: Fill in parameters in orders_report.xlsx and press run*/
/*Purpose: Create Orders report from Running Sales*/
/*OUT: Excel file to sales_report_folder (check configuration.sas)*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\cleanup_xlsx_bak_folder.sas";

%macro read_or_metadata();

	PROC IMPORT OUT=orders_report_md_raw
	            DATAFILE="&or_metadata_file."
	            DBMS=  EXCELCS  REPLACE; 
	RUN;

	data orders_report_md_raw1;
		set orders_report_md_raw;
		if ^missing(coalesceC(of _character_)) or ^missing(coalesce(of _numeric_)) then output;
	run;

	data dmimport.orders_report_md(drop=_: rename=(season=order_season product_form=PF_for_sales_text));
		length hash 8.;
		set orders_report_md_raw1(rename=(season_week_start=_season_week_start));
		hash=1;
		season_week_start=input(_season_week_start, best.);
	run;

%mend read_or_metadata;

%macro orders_report();

	%read_or_metadata();

	proc sql noprint;
		select count(*) into :region trimmed from dmimport.orders_report_md where ^missing(region);
		select count(*) into :country trimmed from dmimport.orders_report_md where ^missing(country);
		select count(*) into :product_line trimmed from dmimport.orders_report_md where ^missing(product_line);
		select count(*) into :species trimmed from dmimport.orders_report_md where ^missing(species);
		select count(*) into :series trimmed from dmimport.orders_report_md where ^missing(series);
		select count(*) into :variety trimmed from dmimport.orders_report_md where ^missing(variety);
		select count(*) into :material trimmed from dmimport.orders_report_md where ^missing(material);
		select count(*) into :PF_for_sales_text trimmed from dmimport.orders_report_md where ^missing(PF_for_sales_text);
		select count(*) into :process_stage trimmed from dmimport.orders_report_md where ^missing(process_stage);
		select count(*) into :mat_div trimmed from dmimport.orders_report_md where ^missing(mat_div);
		select count(*) into :order_season trimmed from dmimport.orders_report_md where ^missing(order_season);
		select count(*) into :product_line_group trimmed from dmimport.orders_report_md where ^missing(product_line_group);
		select count(*) into :season_week_start trimmed from dmimport.orders_report_md where ^missing(season_week_start);
	quit;

	data order_report(
		drop=	rc 
					hash 
					reject 
		keep=	Order_type
					Sls_org
					Sls_off
					Sls_grp
					Soldto_nr
					Soldto_name
					Shipto_nr
					Shipto_name
					Shipto_cntry
					Mat_div
					Mat_grp
					material
					Matdescr
					SchedLine_Cnf_deldte
					Order_week
					Order_month
					Line_crdte
					Ord_rsn
					POnr
					Ord_qty
					Cnf_qty
					Qty_uom
					variety
					Var_descr
					Rsn_rej_cd
					ABC_class
					salrep_nr
					salrep_name
					Product_Line
					species
					series
					PF_for_sales_text
					process_stage
					delivery_week
					delivery_month
					delivery_year
					delivery_season
					delivery_month_season
					region
					territory
					country
					sub_unit
					historical_sales
					actual_sales
					season_week_start
					season_week_end
					product_line_group
					sls_doc_nr
					line_nr
					Itm_Net_val
					Net_value_curr
					);
		length rc hash reject 8.;
		set dmproc.orders_all;
		reject=0;

		%if &region. > 0 %then %do;
			if _n_=1 then do;
				declare hash h_region(dataset: 'dmimport.orders_report_md(where=(^missing(region)))');
					rc=h_region.DefineKey ('region');
					rc=h_region.DefineData ('hash');
					rc=h_region.DefineDone();
			end;
			hash=0;
			rc=h_region.find();
			if hash=0 then reject=1;
		%end;

		%if &country. > 0 %then %do;
			if _n_=1 then do;
				declare hash h_country(dataset: 'dmimport.orders_report_md(where=(^missing(country)))');
					rc=h_country.DefineKey ('country');
					rc=h_country.DefineData ('hash');
					rc=h_country.DefineDone();
			end;
			hash=0;
			rc=h_country.find();
			if hash=0 then reject=1;
		%end;

		%if &product_line. > 0 %then %do;
			if _n_=1 then do;
				declare hash h_product_line(dataset: 'dmimport.orders_report_md(where=(^missing(product_line)))');
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
				declare hash h_species(dataset: 'dmimport.orders_report_md(where=(^missing(species)))');
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
				declare hash h_series(dataset: 'dmimport.orders_report_md(where=(^missing(series)))');
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
				declare hash h_variety(dataset: 'dmimport.orders_report_md(where=(^missing(variety)))');
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
				declare hash h_material(dataset: 'dmimport.orders_report_md(where=(^missing(material)))');
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
				declare hash h_product_form(dataset: 'dmimport.orders_report_md(where=(^missing(PF_for_sales_text)))');
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
				declare hash h_process_stage(dataset: 'dmimport.orders_report_md(where=(^missing(process_stage)))');
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
				declare hash h_mat_div(dataset: 'dmimport.orders_report_md(where=(^missing(mat_div)))');
					rc=h_mat_div.DefineKey ('mat_div');
					rc=h_mat_div.DefineData ('hash');
					rc=h_mat_div.DefineDone();
			end;
			hash=0;
			rc=h_mat_div.find();
			if hash=0 then reject=1;
		%end;

		%if &order_season. > 0 %then %do;
			if _n_=1 then do;
				declare hash h_order_season(dataset: 'dmimport.orders_report_md(where=(^missing(order_season)))');
					rc=h_order_season.DefineKey ('order_season');
					rc=h_order_season.DefineData ('hash');
					rc=h_order_season.DefineDone();
			end;
			hash=0;
			rc=h_order_season.find();
			if hash=0 then reject=1;
		%end;

		%if &product_line_group. > 0 %then %do;
			if _n_=1 then do;
				declare hash h_product_line_group(dataset: 'dmimport.orders_report_md(where=(^missing(product_line_group)))');
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
				declare hash h_season_week_start(dataset: 'dmimport.orders_report_md(where=(^missing(season_week_start)))');
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

	data order_report_rename;
		set order_report;
		label Order_type='Order_type';
		label Sls_doc_nr='Sls_doc_nr';
		label Sls_org='Sls_org';
		label Sls_off='Sls_off';
		label Sls_grp='Sls_grp';
		label Soldto_nr='Soldto_nr';
		label Soldto_name='Soldto_name';
		label Shipto_nr='Shipto_nr';
		label Shipto_name='Shipto_name';
		label Shipto_cntry='Shipto_cntry';
		label Mat_div='Mat_div';
		label Matdescr='Mat_descr';
		label Mat_grp='Mat_grp';
		label Qty_uom='Qty_uom';
		label Var_descr='Var_descr';
		label Rsn_rej_cd='Rsn_rej_cd';
		label ABC_class='ABC_class';
		label SchedLine_Cnf_deldte='SchedLine_Cnf_deldte';
		label Line_crdte='Line_crdte';
		label material='Mat_nr';
		label variety='Var_nr';
		label Ord_rsn='Ord_rsn';
		label POnr='POnr';
		label Ord_qty='Ord_qty';
		label Cnf_qty='Cnf_qty';
		label Itm_Net_val='Itm_Net_val';
		label	Net_value_curr='Net_value_curr';
		label salrep_nr='Salrep_nr';
		label line_nr='Line_nr';
		label salrep_name='Salrep_name';
		label region='Region';
		label territory='Territory';
		label country='Country';
		label sub_unit='Sub_unit';
		label PF_for_sales_text='Product_form';
		label process_stage='Process_stage';
		label historical_sales='Historical_sales';
		label actual_sales='Actual_sales';
		label season_week_start='Season_start_wk';
		label season_week_end='Season_end_wk';
		label order_week='Del_wk';
		label product_line_group='Product_line_group';
		label Product_Line='Product_Line';
		label species='Species';
		label series='Series';
		label delivery_week='Del_wk_YYYYWW';
		label delivery_season='Del_wk_season';
		label delivery_year='Del_year';
		label delivery_month='Del_mm_YYYYMM';
		label delivery_month_season='Del_mm_season';
		label order_month='Del_mm';
	run;

	data order_report_reorder;
	retain 
		Order_type
		Sls_doc_nr
		line_nr
		POnr
		Ord_rsn
		Rsn_rej_cd
		Line_crdte
		SchedLine_Cnf_deldte
		order_week
		order_month
		delivery_year
		Sls_org
		Sls_off
		Sls_grp
		Soldto_nr
		Soldto_name
		Shipto_nr
		Shipto_name
		Shipto_cntry
		ABC_class
		salrep_nr
		salrep_name
		region
		territory
		country
		Product_Line
		product_line_group
		species
		series
		variety
		Var_descr
		material
		Matdescr
		Mat_div
		Mat_grp
		PF_for_sales_text
		process_stage
		season_week_start
		season_week_end
		delivery_month_season
		delivery_season
		delivery_month
		delivery_week
		Ord_qty
		Cnf_qty
		sub_unit
		Qty_uom
		Itm_Net_val
		Net_value_curr
		historical_sales
		actual_sales
		br_actual_sales
		br_historical_sales;
		set order_report_rename;
	run;

	data _null_;
		order_report_file=catx('_', compress(put(today(),yymmdd10.),,'kd'), compress(put(time(), time8.),,'kd'));
		call symput('order_report_file', strip(order_report_file));
	run;

	x "del &sales_report_folder.\Sales_report_&order_report_file..xlsx"; 

	proc export 
	  data=order_report_reorder 
	  dbms=xlsx 
	  outfile="&sales_report_folder.\Sales_report_&order_report_file..xlsx" replace label;
		sheet="Created_on_&order_report_file.";
	run;

	proc export 
	  data=orders_report_md_raw1
	  dbms=xlsx 
	  outfile="&sales_report_folder.\Sales_report_&order_report_file..xlsx";
		sheet="Variant";
	run;

	%cleanup_xlsx_bak_folder(cleanup_folder=%str(&sales_report_folder.\));

%mend;

%orders_report();