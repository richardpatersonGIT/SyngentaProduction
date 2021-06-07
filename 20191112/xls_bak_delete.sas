%macro xlsx_bak_delete(folder=);

	data _null_;
		fname = 'todelete';
		rc = filename(fname, "&file..xlsx.bak");
		rc = fdelete(fname);
		rc = filename(fname);
	run;
	
%mend xlsx_bak_delete;