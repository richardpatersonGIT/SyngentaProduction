/***********************************************************************/
/*Type: Import*/
/*Use: Used as standalone program (press run).*/
/*Purpose: Used in extrapolation. 
           Imports Sales history aka. "Monika files" (weekly download from SAP) from folder (Sales_history_folder=) with name in form (YYYY_WkWW_*.xls*)*/
/*         If those files are already imported into shw(sas library), they are skipped. If you want to overwrite them - delete shw.shw_* sas dataset from shw library*/
/*IN: Sales_weekly_his(multiple files), fist sheet in Excel*/
/*OUT: shw.raw_YYYY_WkWW_* (all columns)*/
/*     shw.shw_YYYY_WkWW_* (essential columns)*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";

%macro import_sales_history(Sales_history_folder=);

  Filename filelist pipe "dir /b /s &Sales_history_folder.\*.xls*"; 
                                                                                   
 Data filelist;                                        
   Infile filelist truncover;
   Input filename $200.;
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

    proc contents data=shw.&dsrawname. out=contents noprint;
    run;

    data shw.&dsname.(keep=sls_org sls_off shipto_cntry matnr var_nr SchedLine_Cnf_deldte cnf_qty order_type mat_div rsn_rej_cd soldto_nr 
                      rename=(matnr=material var_nr=variety));
      length sls_org $4.
             sls_off $4.
             shipto_cntry $2.
             matnr 8.
             var_nr 8.
             SchedLine_Cnf_deldte 8.
             cnf_qty 8.
             order_type $4.
             mat_div $2.
             rsn_rej_cd $2.
             soldto_nr $8.;
      format SchedLine_Cnf_deldte yymmdd10.;
      set shw.&dsrawname.
        (keep=sls_org sls_off shipto_cntry matnr var_nr SchedLine_Cnf_deldte cnf_qty order_type mat_div rsn_rej_cd soldto_nr
         rename=(sls_org=_sls_org
                 sls_off=_sls_off
                 shipto_cntry=_shipto_cntry
                 matnr=_matnr
                 var_nr=_var_nr
                 SchedLine_Cnf_deldte=_SchedLine_Cnf_deldte
                 cnf_qty=_cnf_qty
                 order_type=_order_type
                 mat_div=_mat_div
                 rsn_rej_cd=_rsn_rej_cd
                 soldto_nr=_soldto_nr
        ));

      if vtype(_sls_org)="C" then do;
        sls_org=left(_sls_org);
      end;

      if vtype(_sls_off)="C" then do;
        sls_off=left(_sls_off);
      end;

      if vtype(_shipto_cntry)="C" then do;
        shipto_cntry=left(_shipto_cntry);
      end;

      if vtype(_matnr)="C" then do;
        matnr=input(compress(_matnr,,'kd'), 8.);
      end; else do;
        matnr=_matnr;
      end;

      if vtype(_var_nr)="C" then do;
        var_nr=input(compress(_var_nr,,'kd'), 8.);
      end; else do;
        var_nr=_var_nr;
      end;

      if vtype(_SchedLine_Cnf_deldte)="C" then do;
        SchedLine_Cnf_deldte=input(_SchedLine_Cnf_deldte, 8.)-21916;
      end; else do;
        SchedLine_Cnf_deldte=_SchedLine_Cnf_deldte;
      end;

      if vtype(_cnf_qty)="C" then do;
        if ^missing(_cnf_qty) then do;
          if substr(trim(_cnf_qty), length(trim(_cnf_qty))-2, 1)='.' then do;
            cnf_qty=input(_cnf_qty, comma20.2);  
          end;
          if substr(trim(_cnf_qty), length(trim(_cnf_qty))-2, 1)=',' then do;
            cnf_qty=input(_cnf_qty, commax20.2);  
          end;
        end;
      end; else do;
        cnf_qty=_cnf_qty;
      end;

      if vtype(_order_type)="C" then do;
        order_type=left(_order_type);
      end;

      if vtype(_mat_div)="C" then do;
        mat_div=left(_mat_div);
      end;

      if vtype(_rsn_rej_cd)="C" then do;
        rsn_rej_cd=left(_rsn_rej_cd);
      end; else do;
        if ^missing(_rsn_rej_cd) then do;
          rsn_rej_cd=put(_rsn_rej_cd, 2.);
        end;
      end;

      if vtype(_soldto_nr)="C" then do;
        soldto_nr=left(_soldto_nr);
      end; else do;
        soldto_nr=put(_soldto_nr, 8.);
      end;

      if ^missing(cnf_qty) and ^missing(matnr) then output;
    run;

  %end;

%mend import_sales_history;
  
%import_sales_history(Sales_history_folder=%str(&sales_history_folder.));