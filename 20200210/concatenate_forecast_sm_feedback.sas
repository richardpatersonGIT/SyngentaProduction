/***********************************************************************/
/*Type: Utility*/
/*Use: Used in forecast_report*/
/*Purpose: Concatenate forcast reports (sm feedback, and sm assumption columns) from folder ()*/
/*IN: forecast_reports from each country, sheet='Variety level fcst'*/
/*OUT: work.forecast_sm_feedback_demand*/
/*     work.forecast_sm_feedback_assm*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";

%macro read_forecast_file(forecast_file=, forecast_sheet=);

  PROC IMPORT OUT=forecast_file_raw
              DATAFILE="&forecast_file."
              DBMS=  EXCELCS  REPLACE;
              range="&forecast_sheet.$A8:AX"n;
  RUN;

  data FF_cols(keep=memname name sourcename sourcetype);
    set sashelp.vcolumn;
    length sourcename $32. sourcetype $1.;
    if libname='WORK' and memname='FORECAST_FILE_RAW' then do;
      if type='num' then sourcetype='N';
      if type='char' then sourcetype='C';
      sourcename=upcase(strip(name));
      output;
    end;
  run;

  data FF_cols1;
    length colrename $32. targettype $1.;
    set FF_cols;
    if index(sourcename, 'SM')=1 then do;
      colrename='SM_DEMAND_'||substr(strip(compress(sourcename, '', 'kd')), 1, 4);
      targettype='N';
    end;
    if index(sourcename, 'DM')=1 then do;
      colrename='DMPM_DEMAND_'||substr(strip(compress(sourcename, '', 'kd')), 1, 4);
      targettype='N';
    end;
    if index(sourcename, 'UNDERLYING')=1 then do;
      colrename='ASSUMPTIONS_'||substr(strip(compress(sourcename, '', 'kd')), 1, 4);
      targettype='C';
    end;
    if sourcename='SUM' then do;
      colrename='ORDER';
      targettype='N';
    end;
    if sourcename='COUNTRY' then do;
      colrename='COUNTRY';
      targettype='C';
    end;
    if sourcename='TERRITORY' then do;
      colrename='TERRITORY';
      targettype='C';
    end;
    if sourcename='VARIETY' then do;
      colrename='VARIETY';
      targettype='N';
    end;
    if sourcename='PLC' then do;
      colrename='PLC';
      targettype='C';
    end;
    if sourcename='FUTURE_PLC' then do;
      colrename='FUTURE_PLC';
      targettype='C';
    end;
    if sourcename='VALID_FROM_DATE' then do;
      colrename='VALID_FROM_DATE';
      targettype='N';
    end;
    if sourcename='GLOBAL_PLC' then do;
      colrename='GLOBAL_PLC';
      targettype='C';
    end;
    if sourcename='FUTURE_GLOBAL_PLC' then do;
      colrename='GLOBAL_FUTURE_PLC';
      targettype='C';
    end;
    if sourcename='GLOBAL_VALID_FROM_DATE' then do;
      colrename='GLOBAL_VALID_FROM_DATE';
      targettype='N';
    end;
    if ^missing(colrename) then output;
  run;

  proc sql noprint;
    select 
      strip(sourcename)||'=_'||strip(colrename), 
      strip(colrename), 
      strip(sourcetype)||strip(targettype), 
      strip(colrename),
      strip(sourcename),
      count(*)
    into 
      :FF_rename separated by ' ',
      :FF_keep separated by ' ',
      :FF_cols_conversion separated by '#', 
      :FF_cols_rename separated by '#',
      :FF_cols_name separated by '#',
      :FF_cnt
    from FF_cols1;
  quit;

  data forecast_file(keep=&FF_keep.);
    length VARIETY VALID_FROM_DATE GLOBAL_VALID_FROM_DATE 8.;
    length COUNTRY $10.;
    format VALID_FROM_DATE GLOBAL_VALID_FROM_DATE ddmmyy10.;
    set FORECAST_FILE_RAW(rename=(&FF_rename.));
    %do FF_col=1 %to &FF_cnt.;
      %let FF_col_conversion=%scan(&FF_cols_conversion., &FF_col., '#');
      %let FF_col_rename=%scan(&FF_cols_rename., &FF_col., '#');
      %let FF_col_name=%scan(&FF_cols_name., &FF_col., '#');

      %if "&FF_col_rename."="COUNTRY" or 
          "&FF_col_rename."="TERRITORY" or 
          "&FF_col_rename."="PLC" or
          "&FF_col_rename."="FUTURE_PLC" or 
          "&FF_col_rename."="GLOBAL_PLC" or
          "&FF_col_rename."="GLOBAL_FUTURE_PLC" or
          %index(&FF_col_rename., ASSUMPTIONS)>0 
      %then %do;
        %if "&FF_col_conversion."="NC" %then &FF_col_rename.=strip(put(_&FF_col_rename., best.));;
        %if "&FF_col_conversion."="CC" %then &FF_col_rename.=_&FF_col_rename.;;
      %end;

      %if "&FF_col_rename."="VARIETY" or
          "&FF_col_rename."="ORDER" or
          %index(&FF_col_rename., SM)>0 or
          %index(&FF_col_rename., DMPM)>0
      %then %do;
        %if "&FF_col_conversion."="CN" %then &FF_col_rename.=input(strip(compress(_&FF_col_rename., '', 'kd')), best.);;
        %if "&FF_col_conversion."="NN" %then &FF_col_rename.=_&FF_col_rename.;;
      %end;

      %if "&FF_col_rename."="VALID_FROM_DATE" or
          "&FF_col_rename."="GLOBAL_VALID_FROM_DATE"
      %then %do;
        %if "&FF_col_conversion."="CN" %then if ^missing(_&FF_col_rename.) then &FF_col_rename.=input(_&FF_col_rename., ddmmyy10.);;
        %if "&FF_col_conversion."="NN" %then &FF_col_rename.=_&FF_col_rename.;;
      %end;

    %end;
    if ^missing(variety) then output;
  run;

  proc sort data=forecast_file;
    by COUNTRY VARIETY;
  run;

  proc transpose data=forecast_file out=_FF_sm;
    by COUNTRY VARIETY;
    var SM: ;
  run;

  proc transpose data=forecast_file out=_FF_assumptions;
    by COUNTRY VARIETY;
  var ASSUMPTIONS:;
  run;

  proc transpose data=forecast_file out=_FF_dmpm;
    by COUNTRY VARIETY;
  var DMPM:;
  run;

  data FF_assumptions(drop=_name_);
    set _FF_assumptions(rename=(col1=ASSUMPTION));
    length season 8.;
    SEASON=input(strip(compress(_name_, '', 'kd')), 4.);
  run;

  data FF_sm(drop=_name_);
    set _FF_sm(rename=(col1=SM));
    length season 8.;
    SEASON=input(strip(compress(_name_, '', 'kd')), 4.);
  run;

  data FF_dmpm(drop=_name_);
    set _FF_dmpm(rename=(col1=PMDM));
    length season 8.;
    SEASON=input(strip(compress(_name_, '', 'kd')), 4.);
  run;


  proc sql;
    create table FF_plc as
      select country as REGION, variety, 
             plc, future_plc, valid_from_date,
             global_plc, global_future_plc, global_valid_from_date 
      from forecast_file where order=3;
  quit;

%mend read_forecast_file;

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

    %read_forecast_file(forecast_file=&filename., forecast_sheet=Variety level fcst);

    data feedback_tmp_demand (keep=variety country smf_demand1 smf_demand2 smf_demand3);
      set forecast_file(rename=(
                      sm_demand_&nextseason1.=smf_demand1
                      sm_demand_&nextseason2.=smf_demand2
                      sm_demand_&nextseason3.=smf_demand3));
        smf_demand1=coalesce(smf_demand1, 0);
        smf_demand2=coalesce(smf_demand2, 0);
        smf_demand3=coalesce(smf_demand3, 0);
    run;

    data feedback_tmp_assm (keep=variety country smf_assm1 smf_assm2 smf_assm3);
      set forecast_file(rename=(
                      assumptions_&nextseason1.=smf_assm1
                      assumptions_&nextseason2.=smf_assm2
                      assumptions_&nextseason3.=smf_assm3));
    run;

    %if "&i."="1" %then %do;
      data forecast_sm_feedback_demand;
      retain country variety smf_demand1 smf_demand2 smf_demand3;
        set feedback_tmp_demand;
      run;

      data forecast_sm_feedback_assm;
        length smf_assm1 smf_assm2 smf_assm3 $1000.;
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