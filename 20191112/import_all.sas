/***********************************************************************/
/*Type: Program (press run)*/
/*Use: Run in SAS Enterprise Guide*/
/*Purpose: Imports all data sources*/
/**/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";
%include "&sas_applications_folder.\xls_country_lookup.sas";
%include "&sas_applications_folder.\xls_seasons.sas";
%include "&sas_applications_folder.\xls_fps_assortment.sas";
%include "&sas_applications_folder.\text_pmd_global.sas";
%include "&sas_applications_folder.\text_material_class_table.sas";
%include "&sas_applications_folder.\text_variety_class_table.sas";
%include "&sas_applications_folder.\text_orders.sas";
%include "&sas_applications_folder.\text_delivery_window.sas";
%include "&sas_applications_folder.\xls_bi_seed_assortment.sas";
%include "&sas_applications_folder.\xls_bi_market_seasonality.sas";
%include "&sas_applications_folder.\errors_check.sas";


%macro import_all();

	%let import_start_time=IMPORT STARTED: %sysfunc(date(),worddate.). %sysfunc(time(),timeampm.);

	%IMPORT_SEASONS(seasons_file=%str(&metadata_folder.\seasons_configuration.xlsx));
	%IMPORT_COUNTRY_LOOKUP(cl_file=%str(&metadata_folder.\Country_Lookup.xlsx));
	%IMPORT_FPS_ASSORTMENT(fps_current_folder=%str(&import_folder.\FPS_ASSORTMENT\FPS_ASSORTMENT_CURRENT), fps_deleted_folder=%str(&import_folder.\FPS_ASSORTMENT\FPS_ASSORTMENT_DELETED));
	%IMPORT_PMD_ASSORTMENT(pmd_assortment_folder=%str(&import_folder.\PMD_ASSORTMENT));
	%IMPORT_MATERIAL_CLASS_TABLE(material_class_table_folder=%str(&import_folder.\MATERIAL_CLASS_TABLE));
	%IMPORT_VARIETY_CLASS_TABLE(variety_class_table_folder=%str(&import_folder.\VARIETY_CLASS_TABLE));

	%IMPORT_ORDERS(orders_folder=%str(&import_folder.\ORDERS)); 
	%IMPORT_DELIVERY_WINDOW(delivery_window_folder=%str(&import_folder.\DELIVERY_WINDOW));  
	%IMPORT_BI_SEED_ASSORTMENT(bisa_folder=%str(&import_folder.\BI_SEED_ASSORTMENT));
	%IMPORT_BI_FIXED_SPLIT(bfs_folder=%str(&import_folder.\BI_FIXED_SPLIT)); 

	%PROCESS_PMD_ASSORTMENT();
	%PROCESS_ORDERS();
	%CHECK_ERRORS(); 

	%let import_end_time=IMPORT ENDED: %sysfunc(date(),worddate.). %sysfunc(time(),timeampm.);

	%put &import_start_time.;
	%put &import_end_time.;

%mend import_all;

%import_all();