/***********************************************************************/
/*Type: Utility*/
/*Use: Macro inside a program*/
/*Purpose: (Sas when creating xlsx files, creates also xlsx.bak file) This macro deletes all files with extension ".bak" in given folder*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro cleanup_xlsx_bak_folder(cleanup_folder=);

  %read_folder(folder=%str(&cleanup_folder.), filelist=cleanup_bak_files, ext_mask=bak);

  proc sql noprint;
    select count(*) into :cleanup_bak_files_cnt trimmed from cleanup_bak_files;
  quit;

  %do cl_file=1 %to &cleanup_bak_files_cnt.;

    proc sql noprint;
      select path_file_name into :cleanup_bak_file trimmed from cleanup_bak_files where order = &cl_file.;
    quit;

    data _null_;
      fname = 'todelete';
      rc = filename(fname, "&cleanup_bak_file.");
      rc = fdelete(fname);
      rc = filename(fname);
    run;

  %end;
  
%mend cleanup_xlsx_bak_folder;