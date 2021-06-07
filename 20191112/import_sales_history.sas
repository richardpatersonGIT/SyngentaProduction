/***********************************************************************/
/*Type: Import*/
/*Use: Used as standalone program (press run).*/
/*Purpose: Used in extrapolation. 
           Imports Sales history aka. "Monika files" (weekly download from SAP) from folder (Sales_history_folder=) with name in form (YYYY_WkWW_*.xls*)*/
/*	       If those files are already imported into shw(sas library), they are skipped. If you want to overwrite them - delete shw.shw_* sas dataset from shw library*/
/*IN: Sales_weekly_his(multiple files), fist sheet in Excel*/
/*OUT: shw.raw_YYYY_WkWW_* (all columns)*/
/*     shw.shw_YYYY_WkWW_* (essential columns)*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";

%macro import_sales_history(Sales_history_folder=);

  Filename filelist pipe "dir /b /s &Sales_history_folder.\*.xls*"; 
                                                                                   
 Data filelist;                                        
   Infile filelist truncover;
   Input filename $100.;
 Run; 

	data filelist1;
		set filelist;
		fn=scan(filename, -1, '\');
		dsrawname=cats('RAW_',upcase(tranwrd(scan(fn,1,'.'), ' ', '_')));
		dsname=cats('SHW_',upcase(tranwrd(scan(fn,1,'.'), ' ', '_')));
		type=scan(filename, -2, '\');
		year=input(substr(fn, 1, 4), 4.);
		week=input(compress(substr(fn, 8, 2),, "kd"),  8.);
	run;

	proc sort data=filelist1;
		by dsname;
	run;

	proc contents data=shw._all_ out=shw_contents noprint;
	run;

	proc sql noprint;
		create table shw_contents1 as
		select distinct memname as dsname from shw_contents order by dsname;
	quit;

	data filelist2;
		merge filelist1 (in=fl) shw_contents1 (in=shw);
		by dsname;
		if fl and not shw then output;
	run;

	data filelist3;
		set filelist2;
		order=_n_;
	run;

	proc sql noprint;
		select count(*) into : flcnt from filelist3;
	quit;
	
	%do i=1 %to &flcnt.;
		proc sql noprint;
			select filename, fn, dsrawname, dsname, type, year, week into :filename, :fn, :dsrawname, :dsname, :type, :year, :week from filelist3 where order = &i.;
		quit;
		%let filename=&filename.;
		%let fn=&fn.;
		%let dsrawname=&dsrawname.;
		%let dsname=&dsname.;
		%let type=&type.;
		%let year=&year.;
		%let week=&week.;
		%put &=filename, &=fn, &=type, &=year, &=week;

		PROC IMPORT OUT=shw.&dsrawname.
								DATAFILE= "&filename." 
		            DBMS=xlsx REPLACE;
		     				GETNAMES=YES;
		RUN;

		/*variables fix*/
		proc contents data=shw.&dsrawname. out=contents noprint;
		run;

		proc sql noprint;
			select type into :schedline_cnf_deldte_type trimmed from contents where lower(name)="schedline_cnf_deldte";
			select type into :cnf_qty_type trimmed from contents where lower(name)="cnf_qty";
			select type into :matnr_type trimmed from contents where lower(name)="matnr";
			select type into :rsn_rej_cd_type trimmed from contents where lower(name)="rsn_rej_cd";
			select type into :var_nr_type trimmed from contents where lower(name)="var_nr";
		quit;

		data shw.&dsname.(keep=sls_org sls_off shipto_cntry matnr var_nr SchedLine_Cnf_deldte cnf_qty order_type mat_div rsn_rej_cd rename=(matnr=material var_nr=variety));
			set shw.&dsrawname.
				(rename=(
									mat_div=_mat_div
									order_type=_order_type
									shipto_cntry=_shipto_cntry
									sls_off=_sls_off
									sls_org=_sls_org
									%if "&schedline_cnf_deldte_type."="2" %then %do;
										SchedLine_Cnf_deldte=_SchedLine_Cnf_deldte
									%end;
									%if "&cnf_qty_type."="2" %then %do;
										cnf_qty=_cnf_qty
									%end;
									%if "&matnr_type."="2" %then %do;
										matnr=_matnr
									%end;
									%if "&rsn_rej_cd_type."="2" %then %do;
										rsn_rej_cd=_rsn_rej_cd
									%end;
									%if "&var_nr_type."="2" %then %do;
										var_nr=_var_nr
									%end;
				));
				length mat_div $2.
				order_type $4.
				;
			%if "&schedline_cnf_deldte_type."="2" %then %do;
				SchedLine_Cnf_deldte=input(_SchedLine_Cnf_deldte, 8.);
				SchedLine_Cnf_deldte=SchedLine_Cnf_deldte-21916;	
			%end;
			%if "&cnf_qty_type."="2" %then %do;
/*				_cnf_qty=compress(_cnf_qty,,'kd');*/
/*				if ^missing(_cnf_qty) then cnf_qty=input(_cnf_qty, 8.);*/
				if ^missing(_cnf_qty) then do;
					if substr(trim(_cnf_qty), length(trim(_cnf_qty))-2, 1)='.' then do;
						cnf_qty=input(_cnf_qty, comma20.2);	
					end;
					if substr(trim(_cnf_qty), length(trim(_cnf_qty))-2, 1)=',' then do;
						cnf_qty=input(_cnf_qty, commax20.2);	
					end;
				end;
			%end;
			%if "&matnr_type."="2" %then %do;
				_matnr=compress(_matnr,,'kd');
				if ^missing(_matnr) then matnr=input(_matnr, 8.);
			%end;
			%if "&rsn_rej_cd_type."="2" %then %do;
				_rsn_rej_cd=compress(_rsn_rej_cd,,'kd');
				if ^missing(_rsn_rej_cd) then rsn_rej_cd=input(_rsn_rej_cd, 8.);
			%end;
			%if "&var_nr_type."="2" %then %do;
				_var_nr=compress(_var_nr,,'kd');
				if ^missing(_var_nr) then var_nr=input(_var_nr, 8.);
			%end;
			if strip(_mat_div) in ('6A', '6B', '6C') then mat_div=strip(_mat_div);
			if length(strip(_order_type))=4 then order_type=strip(_order_type);
			if length(strip(_shipto_cntry))=2 then shipto_cntry=strip(_shipto_cntry);
			if length(strip(_sls_off))=4 then sls_off=strip(_sls_off);
			if length(strip(_sls_org))=4 then sls_org=strip(_sls_org);
			if ^missing(cnf_qty) and ^missing(matnr) then output;
		run;
		/*variables fix*/
	%end;

%mend import_sales_history;
	
%import_sales_history(Sales_history_folder=%str(&sales_history_folder.));