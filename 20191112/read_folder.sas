/***********************************************************************/
/*Type: Utility*/
/*Use: Macro inside a program*/
/*Purpose: Create a sas dataset (filelist=) with filenames of the files in specified folder (folder=) with the proper extension_mask (ext_mask)*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";

%macro read_folder(folder=, filelist=, ext_mask=);
  Filename filelist pipe "dir /b ""&folder.\*"""; 
                                                                                   
 Data filelist_raw;                                        
   Infile filelist truncover;
   Input filename $100.;
 Run; 

	data filelist(drop=filename);
		length path_file_name path_folder file_name_full file_name file_name_ext ext_mask $200.;
		set filelist_raw;
		file_name_full=scan(filename, -1, '\');
		file_name=scan(file_name_full, 1, '.');
		file_name_ext=scan(file_name_full, -1, '.');
		path_folder="&folder.\";
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