/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %import_bi_seed_assortment(...)*/
/*Purpose: Imports BI Seed Assortment from folder (bisa_folder=) with .xlsx extension*/
/*IN: bi_seed_assortment(1 file), extension=xlsx, sheet='BI Assortment'*/
/*OUT: dmimport.bi_seed_assortment*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro import_bi_seed_assortment(bisa_folder=);

	%read_folder(folder=%str(&bisa_folder.), filelist=bisa_files, ext_mask=xlsx);

	proc sql noprint;
	select path_file_name into :bisa_file trimmed from bisa_files where order=1;
	quit;

	PROC IMPORT OUT=bi_seed_assortment_raw
	            DATAFILE="&bisa_file."
	            DBMS=  EXCELCS   REPLACE;
							sheet='BI Assortment';
	RUN;

	data dmimport.bi_seed_assortment;
		set bi_seed_assortment_raw(rename=(variety_number=variety bulk=material_bulk));
	run;

%mend import_bi_seed_assortment;