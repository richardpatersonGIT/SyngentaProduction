/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %IMPORT_DELIVERY_WINDOW(...)*/
/*Purpose: Imports Delivery window from folder (delivery_window_folder=) with .txt extension*/
/*IN: Delivery_window(1 file), extension=txt, delimiter=;*/
/*OUT: dmimport.delivery_window*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro IMPORT_DELIVERY_WINDOW(delivery_window_folder=);

	%read_folder(folder=%str(&delivery_window_folder.), filelist=delivery_window, ext_mask=txt);

	proc sql noprint;
	select path_file_name into :dw_file trimmed from delivery_window where order=1;
	quit;

data delivery_window_raw;
infile "&dw_file." delimiter = ';' MISSOVER DSD lrecl=32767  firstobs=2;
	length 	Species $4.
					Species_Desc $29.
					_Variety $8.
					Variety_Desc $30.
					_Material $8.
					Material_Description $40.
					Plant $4.
					Process_stage $3.
					Process_stage_Desc $20.
					PLC_code $2.
					Material_Group $6.
					Delivery_Week $6.
					Order_Week $6.
					URC_Week $6.
					Qty_in_Box $7.
					Qty_in_Pc $11.
					Open_Closed_week_status $4.
					Requested_EAME_Sales $3.
					Confirmed_EAME_Sales $3.
					Unconfirmed_EAME_Sales $1.
					Constrained_Demand $1.;

	input Species
				Species_Desc
				_Variety
				Variety_Desc
				_Material
				Material_Description
				Plant
				Process_stage
				Process_stage_Desc
				PLC_code
				Material_Group
				Delivery_Week
				Order_Week
				URC_Week
				Qty_in_Box
				Qty_in_Pc
				Open_Closed_week_status
				Requested_EAME_Sales
				Confirmed_EAME_Sales
				Unconfirmed_EAME_Sales
				Constrained_Demand
				;
	run;

	data dmimport.delivery_window(keep=variety material delivery_week_year delivery_week_week status);
		set delivery_window_raw;
		length delivery_week_year delivery_week_week 8.;
		material=input(_material, 8.);
		variety=input(_variety, 8.);
		delivery_week_year=input(substr(put(strip(delivery_week), 6.), 1, 4), 8.);
		delivery_week_week=input(substr(put(strip(delivery_week), 6.), 5, 2), 8.);
		if open_closed_week_status = 'CLSD' then status=0;
		if open_closed_week_status = 'APPR' then status=1;
	run;

%mend IMPORT_DELIVERY_WINDOW;