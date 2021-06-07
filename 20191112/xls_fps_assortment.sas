/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %import_fps_assortment(...)*/
/*Purpose: Imports FPS Assortment from folders [current-(fps_current_folder=)] [deleted-(fps_deleted_folder=)] with .xlsx extension*/
/*IN: fps_assortment_current and fps_assortment_deleted(2 files), extension=xls*, sheet='list'*/
/*OUT: dmimport.FPS_Assortment*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro import_fps_assortment(fps_current_folder=, fps_deleted_folder=);

	%read_folder(folder=%str(&fps_current_folder.), filelist=fps_current, ext_mask=xls);
	%read_folder(folder=%str(&fps_deleted_folder.), filelist=fps_deleted, ext_mask=xls);

	proc sql noprint;
	select path_file_name into :fps_current_file trimmed from fps_current where order=1;
	select path_file_name into :fps_deleted_file trimmed from fps_deleted where order=1;
	quit;

	PROC IMPORT OUT=FPS_Assortment_current_raw 
	            DATAFILE="&fps_current_file."
	            DBMS=  EXCELCS   REPLACE;
	            SHEET='list'; 
	RUN;

	PROC IMPORT OUT=FPS_Assortment_deleted_raw 
	            DATAFILE="&fps_deleted_file."
	            DBMS=  EXCELCS   REPLACE;
	            SHEET='list'; 
	RUN;

	data FPS_Assortment_current(drop=_:);
			set FPS_Assortment_current_raw(keep= Material_ID	Material_basic_description__EN_	Material_sales_text__EN_	Material_division	PF_for_sales_text	Sales_unit	Alt__Unit	Sub_unit	DM_Process_stage	FND_Variety_Nb	PMD_Variety_Nb	FPS_variety_name	Product_line_Reporting	Forecasting_group	Species_name	FPS_series_description	Curr_Var_PLC_FPS	Curr_Mat_PLC replaced_by
																	rename=(		Material_ID=_Material
																							Material_basic_description__EN_=Material_basic_description_EN
																							Material_sales_text__EN_=Material_sales_text_EN
																							Alt__Unit=Alt_Unit
																							FND_Variety_Nb=_Variety
																						));
		filename='FPS_current';
		hash_product_line=lowcase(product_line_reporting); /*lower case for joining*/
		hash_species_name=lowcase(species_name); /*lower case for joining*/
		material=input(strip(_material),8.);
		Variety=input(strip(_Variety),8.);
		if missing(sub_unit) then do;
			sub_unit=1;
			if material_division='6A' then bulk_6A=1;
		end;
		if missing(pf_for_sales_text) then pf_for_sales_text=process_stage_differentiator;/*for QT2*/
		region='FPS';
		if ^missing(material) then output;
	run;

	data FPS_Assortment_deleted(drop=_:);
		set FPS_Assortment_deleted_raw(keep= Material_ID	Material_basic_description__EN_	Material_sales_text__EN_	Material_division	PF_for_sales_text	Sales_unit	Alt__Unit	Sub_unit	DM_Process_stage	FND_Variety_Nb	PMD_Variety_Nb	FPS_variety_name	Product_line_Reporting	Forecasting_group	Species_name	FPS_series_description	Curr_Var_PLC_FPS	Curr_Mat_PLC replaced_by
																	rename=(		Material_ID=_Material
																							Material_basic_description__EN_=Material_basic_description_EN
																							Material_sales_text__EN_=Material_sales_text_EN
																							Alt__Unit=Alt_Unit
																							FND_Variety_Nb=_Variety
																						));
		filename='FPS_deleted';
		hash_product_line=lowcase(product_line_reporting); /*lower case for joining*/
		hash_species_name=lowcase(species_name); /*lower case for joining*/
		material=input(strip(_material),8.);
		Variety=input(strip(_Variety),8.);
		if missing(sub_unit) then do;
			sub_unit=1;
			if material_division='6A' then bulk_6A=1;
		end;
		if missing(pf_for_sales_text) then pf_for_sales_text=process_stage_differentiator;/*for QT2*/
		region='FPS';
		if ^missing(material) then output;
	run;

	proc sql noprint;
		delete from FPS_Assortment_deleted where material in (select distinct material from FPS_Assortment_current);
	quit;

	data FPS_Assortment1;
		length FPS_series_description $25.;
		set FPS_Assortment_current FPS_Assortment_deleted;	
	run;

	proc sort data=FPS_Assortment1 out=FPS_Assortment1_sorted dupout=FPS_Assortment1_dup nodupkey;
	by material;
	run;

	data dmimport.FPS_Assortment;
		set FPS_Assortment1_sorted;	
	run;

%mend import_fps_assortment;