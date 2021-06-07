/***********************************************************************/
/*Type: Utility*/
/*Use: Macro inside a program*/
/*Purpose: Create a sas dataset (filelist=) with filenames of the files in specified folder (folder=) with the proper extension_mask (ext_mask), if parameter (subfolders=) is set to Y program will look also into subfolders*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";

%macro read_folder(folder=, filelist=, ext_mask=, subfolders=N);

  %if "&subfolders."="Y" %then %do;
    Filename filelist pipe "dir /b /s ""&folder.\*"""; 
  %end; %else %do;
    Filename filelist pipe "dir /b ""&folder.\*""";
  %end;
                                                                              
  Data filelist_raw;                                        
   Infile filelist truncover;
   Input filename $400.;
  Run; 

  data filelist(drop=filename);
    length path_file_name path_folder file_name_full file_name file_name_ext ext_mask $200.;
    set filelist_raw;
    file_name_full=scan(filename, -1, '\');
    file_name=scan(file_name_full, 1, '.');
    file_name_ext=scan(file_name_full, -1, '.');
    %if "&subfolders."="Y" %then %do;
      path_folder=substr(filename, 1, length(filename)-length(scan(filename, -1, '\')));
    %end; %else %do;
      path_folder="&folder.\";
    %end;
    path_file_name=cats(path_folder,file_name_full);
    ext_mask="&ext_mask.";
  run;

  data filelist1;
    set filelist;
    if missing(strip(ext_mask)) then do;
      output;
    end; else do;
      if index(lowcase(strip(file_name_ext)), lowcase(strip(ext_mask))) > 0 then output;
    end;
  run;

  data &filelist.;
    set filelist1;
    order=_n_;
  run;

%mend read_folder;