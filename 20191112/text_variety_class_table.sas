/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %IMPORT_VARIETY_CLASS_TABLE(...)*/
/*Purpose: Imports Variety Class Table from folder (variety_class_table_folder=) with .txt extension*/
/*IN: Variety_class_table(multiple files), extension=txt, delimiter=tab*/
/*OUT: dmimport.variety_class_table*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro IMPORT_VARIETY_CLASS_TABLE(variety_class_table_folder=);

	%read_folder(folder=%str(&variety_class_table_folder.), filelist=variety_class_files, ext_mask=txt);

	proc datasets lib=work nolist;
		delete variety_class_table_all;
	run;

	proc sql noprint;
		select count(*) into :variety_class_files_cnt from variety_class_files;
	quit;

	%do mc_file=1 %to &variety_class_files_cnt.;

		proc sql noprint;
			select path_file_name into :variety_class_file trimmed from variety_class_files where order = &mc_file.;
		quit;

		data variety_class_table_raw;
			infile "&variety_class_file." dlm='09'x dsd MISSOVER lrecl=32767  firstobs=4;

			length 
				_Material $18.
				Material_Desc $40.
				Division $2.
				Material_Group $7.
				Cnvt_Unit $3.
				Product_Hierarchy $10.
				Basic_Material $1.
				Material_Type $4.
				Old_Material_Number $1.
				Product_Allocation $1.
				Variety_Number_New $1.
				Variety_Description $1.
				Variety_Number_Old $1.
				Variety_Division $1.
				Class $3.
				Class_Description $7.
				Batch_Class $3.
				Class_Type $1.
				Brand $1.
				Buom $5.
				Caliber $3.
				Generation $5.
				Germ_Level $1.
				Grouping $1.
				Label_Code $1.
				Material_Desc1 $1.
				Material_Nr $1.
				Packing_Type $1.
				Proc_Stage $1.
				_Quantity $1.
				Quom $1.
				Reference_Code $1.
				Species $1.
				Srg $1.
				State_Certif $1.
				Sub_Brand $3.
				Treatment $1.
				Tsw $1.
				Variance_Key $2.
				Variety_Name $1.
				_Variety_Nr $1.;

		input 
			_Material $
			Material_Desc $
			Division $
			Material_Group $
			Cnvt_Unit $
			Product_Hierarchy $
			Basic_Material $
			Material_Type $
			Old_Material_Number $
			Product_Allocation $
			Variety_Number_New $
			Variety_Description $
			Variety_Number_Old $
			Variety_Division $
			Class $
			Class_Description $
			Batch_Class $
			Class_Type $
			Brand $
			Buom $
			Caliber $
			Generation $
			Germ_Level $
			Grouping $
			Label_Code $
			Material_Desc1 $
			Material_Nr $
			Packing_Type $
			Proc_Stage $
			_Quantity $
			Quom $
			Reference_Code $
			Species $
			Srg $
			State_Certif $
			Sub_Brand $
			Treatment $
			Tsw $
			Variance_Key $
			Variety_Name $
			_Variety_Nr $
		;
		run;

		data variety_class_table (drop=_: rename=(Material_Desc=Variety_desc));
			set variety_class_table_raw;
			variety=input(substr(_material, 11, 8), 8.);
			region='BI';
		run;

		proc append base=variety_class_table_all data=variety_class_table;
		run;

	%end;

	proc sort data=variety_class_table_all out=variety_class_table_sorted dupout=variety_class_table_dup nodupkey;
		by variety;
	run;

	data dmimport.variety_class_table;
		set variety_class_table_sorted;
	run;

%mend IMPORT_VARIETY_CLASS_TABLE;