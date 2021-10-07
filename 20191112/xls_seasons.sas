/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %IMPORT_SEASONS(...)*/
/*Purpose: Imports seasons and product_line_group mappings from excel file (seasons_file=)*/
/*IN: Seasons(1 excel file) sheet='seasons'*/
/*OUT: dmimport.seasons_general - general rules*/
/*     dmimport.seasons_grouping_code_exc - exceptions based on grouping code*/
/*     dmimport.seasons_species_exc - exceptions based on species name*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";

%MACRO IMPORT_SEASONS(seasons_file=);

	PROC IMPORT OUT=seasons_raw DATAFILE= "&seasons_file." 
	            DBMS=xlsx REPLACE;
	     SHEET="seasons"; 
	     GETNAMES=YES;
	RUN;

	data dmimport.seasons_general(drop=species_name hash_species_name grouping_code hash_grouping_code) 
			 dmimport.seasons_grouping_code_exc(drop=species_name hash_species_name)
			 dmimport.seasons_species_exc(drop=grouping_code hash_grouping_code);
		length plg_rule $10.;
		set seasons_raw;
		hash_product_line=lowcase(strip(product_line));
		hash_species_name=lowcase(strip(species_name));
		hash_grouping_code=lowcase(strip(grouping_code));
		if missing(species_name) and missing(grouping_code) then output dmimport.seasons_general;
		if ^missing(grouping_code) then output dmimport.seasons_grouping_code_exc;
		if ^missing(species_name) then output dmimport.seasons_species_exc;
	run;

%MEND IMPORT_SEASONS;