/***********************************************************************/
/*Type: Utility*/
/*Use: Used in forecast_report*/
/*Purpose: Concatenate forcast reports (sm feedback, and sm assumption columns) from folder ()*/
/*IN: forecast_reports from each country, sheet='Variety level fcst'*/
/*OUT: work.forecast_sm_feedback_demand*/
/*     work.forecast_sm_feedback_assm*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";

%macro concatenate_forecast_sm_feedback(sm_feedback_folder=);

	%let sm_feedback_folder=&sm_feedback_folder.;
	Filename filelist pipe "dir /b /s ""&sm_feedback_folder.\*.xls*"""; 
	                                                                             
	Data filelist;      
	 	length filename $200.;
		Infile filelist truncover;
		Input _filename $200.;
		filename=strip(_filename);
	Run; 

	data filelist1;
		set filelist;
		fn=scan(filename, -1, '\');
		dsrawname=cats('RAW_',upcase(tranwrd(scan(fn,1,'.'), ' ', '_')));
		dsname=cats('SHW_',upcase(tranwrd(scan(fn,1,'.'), ' ', '_')));
		type=scan(filename, -2, '\');
		year=input(substr(fn, 1, 4), 4.);
		week=input(compress(substr(fn, 8, 2),, "kd"),  8.);
		order=_n_;
	run;

	proc sql noprint;
		select count(*) into :flcnt from filelist1;
	quit;

	%do i=1 %to &flcnt.;

		proc sql noprint;
			select filename into :filename from filelist1 where order=&i.;
		quit;

		PROC IMPORT OUT=feedback_tmp
		          DATAFILE="&filename."
		          DBMS=  EXCELCS   REPLACE;
		          SHEET="Variety level fcst"; 
		RUN;

		proc contents data=feedback_tmp out=feedback_contents noprint;
		run;

	data sm_feedback_cols;
		length varnum 8. columnname $32.;
		varnum=4; /*Excel column D*/
		columnname="Country"; 
		output;
		varnum=8; /*Excel column H*/
		columnname="Variety";
		output;
		varnum=31; /*Excel column AE*/
		columnname="smf_demand1";
		output;
		varnum=35; /*Excel column AI*/
		columnname="smf_demand2";
		output;
		varnum=39; /*Excel column AM*/
		columnname="smf_demand3";
		output;
		varnum=33; /*Excel column AG*/
		columnname="smf_assm1";
		output;
		varnum=37; /*Excel column AK*/
		columnname="smf_assm2";
		output;
		varnum=41; /*Excel column AO*/
		columnname="smf_assm3";
		output;
	run;

		proc sql noprint;
			select compress(name||'=_'||columnname) into :renamestring separated by ' ' from sm_feedback_cols fcols
			left join feedback_contents fcnts on fcols.varnum=fcnts.varnum;
		quit;

		data feedback_tmp_demand (keep=variety country smf_demand1 smf_demand2 smf_demand3);
			set feedback_tmp(rename=(&renamestring.));
			length country $6. variety smf_demand1 smf_demand2 smf_demand3 8.;
			country=strip(_country);
			variety=input(strip(_variety), 8.);
			smf_demand1=input(_smf_demand1, comma20.);
			smf_demand2=input(_smf_demand2, comma20.);
			smf_demand3=input(_smf_demand3, comma20.);
			if ^missing(Variety) then output;
		run;

		data feedback_tmp_assm (keep=variety country smf_assm1 smf_assm2 smf_assm3);
			set feedback_tmp(rename=(&renamestring.));
			length country $6. variety 8. smf_assm1 smf_assm2 smf_assm3 $1000.;
			country=strip(_country);
			variety=input(strip(_variety), 8.);
			smf_assm1=strip(_smf_assm1);
			smf_assm2=strip(_smf_assm2);
			smf_assm3=strip(_smf_assm3);
			if ^missing(Variety) then output;
		run;

		%if "&i."="1" %then %do;
			data forecast_sm_feedback_demand;
				set feedback_tmp_demand;
			run;

			data forecast_sm_feedback_assm;
				set feedback_tmp_assm;
			run;
		%end; %else %do;
			proc append base=forecast_sm_feedback_demand data=feedback_tmp_demand force;
			run;

			proc append base=forecast_sm_feedback_assm data=feedback_tmp_assm force;
			run;
		%end;

	%end;

%mend concatenate_forecast_sm_feedback;