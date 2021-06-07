/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %IMPORT_MATERIAL_CLASS_TABLE(...)*/
/*Purpose: Imports Material Class Table from folder (material_class_table_folder=) with .txt extension*/
/*IN: Material_class_table(multiple files), extension=txt, delimiter=tab*/
/*OUT: dmimport.Material_class_table*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro IMPORT_MATERIAL_CLASS_TABLE(material_class_table_folder=);

	%read_folder(folder=%str(&material_class_table_folder.), filelist=material_class_files, ext_mask=txt);

	proc datasets lib=work nolist;
		delete Material_class_table_all;
	run;

	proc sql noprint;
		select count(*) into :material_class_files_cnt from material_class_files;
	quit;

	%do mc_file=1 %to &material_class_files_cnt.;

		proc sql noprint;
			select path_file_name into :material_class_file trimmed from material_class_files where order = &mc_file.;
		quit;

		data Material_class_table_raw;
			infile "&material_class_file." dlm='09'x dsd MISSOVER lrecl=32767  firstobs=4;

			length 
				_Material $18.
				Material_Desc $40.
				Division $2.
				Material_Group $6.
				Cnvt_Unit $2.
				Product_Hierarchy $10.
				Basic_Material $8.
				Material_Type $4.
				Old_Material_Number $18.
				Product_Allocation $1.
				Variety_Number_New $8.
				Variety_Description $40.
				Variety_Number_Old $1.
				Variety_Division $1.
				Class $3.
				Class_Description $5.
				Batch_Class $3.
				Class_Type $18.
				Brand $2.
				Buom $3.
				Caliber $1.
				Generation $2.
				Germ_Level $1.
				Grouping $3.
				Label_Code $1.
				Material_Desc1 $30.
				Material_Nr $8.
				Packing_Type $2.
				Proc_Stage $3.
				_Quantity $3.
				Quom $2.
				Reference_Code $1.
				Species $4.
				Srg $3.
				State_Certif $1.
				Sub_Brand $1.
				Treatment $3.
				Tsw $15.
				Variance_Key $11.
				Variety_Name $30.
				_Variety_Nr $8.;

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

		data Material_class_table (drop=_:);
			set Material_class_table_raw;
			length Material Variety Quantity sub_unit 8.;
			Material=input(substr(_Material, 11, 8), 8.);
			Variety=input(_Variety_nr, 8.);
			Quantity=input(_Quantity, 8.);
			if missing(Quantity) then Quantity=1;
			sub_unit=Quantity;
			if QUOM='SD' then sub_unit=Quantity/1000;
				else sub_unit=Quantity;
			region='BI';
		run;

		proc append base=Material_class_table_all data=Material_class_table;
		run;

	%end;

	proc sort data=Material_class_table_all out=material_class_table_sorted dupout=material_class_table_dup nodupkey;
		by Material;
	run;

	data dmimport.Material_class_table;
		set material_class_table_sorted;
	run;

%mend IMPORT_MATERIAL_CLASS_TABLE;