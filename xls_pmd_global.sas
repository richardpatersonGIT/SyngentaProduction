/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %IMPORT_PMD_ASSORTMENT(...)  %PROCESS_PMD_ASSORTMENT()*/
/*Purpose: Imports (IMPORT_PMD_ASSORTMENT) and process (PROCESS_PMD_ASSORTMENT) PMD Assortment Global files from folder (pmd_assortment_folder=) with .xlsx extension*/
/*IN: pmd_assortment(1 file), extension=xlsx, delimiter = ';'*/
/*    dmimport.seasons_general*/
/*    dmimport.seasons_Grouping_Code_exc*/
/*    dmimport.seasons_species_exc*/
/*OUT: DMPROC.PMD_assortment*/
/***********************************************************************/
OPTION VALIDVARNAME=V7;
%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro IMPORT_PMD_ASSORTMENT(pmd_assortment_folder=);

  %read_folder(folder=%str(&pmd_assortment_folder.), filelist=pmd_assortment_files, ext_mask=xlsx);

  proc sql noprint;
    select count(*) into :pmd_files_cnt from pmd_assortment_files;
  quit;

  %do pmd_file=1 %to &pmd_files_cnt.;

    proc sql noprint;
      select path_file_name into :pmd_file_path trimmed from pmd_assortment_files where order=&pmd_file.;
    quit;

    PROC IMPORT OUT=pmd_raw
                DATAFILE="&pmd_file_path."
                DBMS=  XLSX   REPLACE;
                sheet='PMD Flower Report';
    RUN;

    data pmd_raw_header;
      set pmd_raw(obs=4 firstobs=3);
    run;

    proc transpose data=pmd_raw_header out=pmd_raw_header1(drop=_label_);
      var _all_;
    run;

    data pmd_raw_header2;
      set pmd_raw_header1;
      length 
        rcol1 $6. 
        org_column_name $200.
        org_column_rename $32.;
      retain rcol1;
      
      col2=substr(lowcase(translate(strip(col2), '_____', '() -.')),1, coalesce(length(col2), 24));
      if ^missing(col1) then do;
        rcol1=strip(col1);
      end;
      if ^missing(rcol1) then do;
        org_column_name=strip(rcol1)||'_'||strip(col2);
      end; else do;
        org_column_name=strip(col2);
      end;
      org_column_rename=substr(strip(org_column_name), 1, min(length(org_column_name), 31));
    run;

    proc sql noprint;
      select strip(_name_)||'=_'||strip(org_column_rename) into :rename_string separated by ' ' 
      from pmd_raw_header2;
    quit;

    data pmd_raw1;
      set pmd_raw(rename=(&rename_string.) firstobs=5);
    run;

    data pmd1(drop=_:);
      set pmd_raw1; 
      length 
        product_line	$19.
        multiplication_indicator	$2.
        species_code	$4.
        species	$29.
        global_series_code $3.
        series	$50.
        originator_name	$100.
        genetic_owner	$50.
        owner_group_code	8.
        grouping_code	$1.
        parent_line_indicator	$1.
        short_code	$5.
        variety	8.
        variety_research_code	$50.
        breeder	$50.
        lead_product_manager	$100.
        GLOBAL_variety_name	$100.
        GLOBAL_current_plc	$2.
        GLOBAL_future_plc	$2.
        GLOBAL_date_of_future_plc	8.
        GLOBAL_replace_by	8.
        GLOBAL_replacement_date	8.
        GLOBAL_abc	$3.
        GLOBAL_brand	$2.
        EUROPE_series_code $3.
        EUROPE_series_description $44.
        EUROPE_variety_name	$100.
        EUROPE_current_plc	$2.
        EUROPE_future_plc	$2.
        EUROPE_date_of_future_plc	8.
        EUROPE_replace_by	8.
        EUROPE_replacement_date	8.
        EUROPE_abc	$3.
        EUROPE_channel_eame	$1.
        EUROPE_channel_floranova	$1.
        EUROPE_regional_product_manager	$100.
        NAM_series_code $3.
        NAM_series_description $44.
        NAM_variety_name  $100.
        NAM_current_plc  $2.
        NAM_future_plc  $2.
        NAM_date_of_future_plc  8.
        NAM_replace_by  8.
        NAM_replacement_date  8.
        NAM_abc  $3.
        NAM_regional_product_manager  $100.
        APAC_series_code $3.
        APAC_series_description $44.
        APAC_variety_name	$100.
        APAC_current_plc	$2.
        APAC_future_plc	$2.
        APAC_date_of_future_plc	8.
        APAC_replace_by	8.
        APAC_replacement_date	8.
        APAC_abc	$3.
        APAC_channel_bi	$1.
        APAC_channel_floranova	$1.
        APAC_channel_japan	$1.
        APAC_regional_product_manager	$100.
        AME_series_code $3.
        AME_series_description $44.
        AME_variety_name	$100.
        AME_current_plc	$2.
        AME_future_plc	$2.
        AME_date_of_future_plc	8.
        AME_replace_by	8.
        AME_replacement_date	8.
        AME_abc	$3.
        AME_channel_bi	$1.
        AME_channel_floranova	$1.
        AME_regional_product_manager	$100.
        LATAM_series_code $3.
        LATAM_series_description $44.
        LATAM_variety_name	$100.
        LATAM_current_plc	$2.
        LATAM_future_plc	$2.
        LATAM_date_of_future_plc	8.
        LATAM_replace_by	8.
        LATAM_replacement_date	8.
        LATAM_abc	$3.
        LATAM_channel_bi	$1.
        LATAM_channel_floranova	$1.
        LATAM_regional_product_manager	$100.
        ;
      format 
        GLOBAL_date_of_future_plc 
        GLOBAL_replacement_date
        EUROPE_date_of_future_plc
        EUROPE_replacement_date
        NAM_date_of_future_plc
        NAM_replacement_date
        APAC_date_of_future_plc
        APAC_replacement_date
        AME_date_of_future_plc
        AME_replacement_date
        LATAM_date_of_future_plc
        LATAM_replacement_date
        yymmdd10.;

      /*GENERAL*/
      product_line=strip(_product_line);
      multiplication_indicator=strip(_multiplication_indicator);
      species_code=strip(_species_code);
      species=strip(_species);
      global_series_code=strip(_global_series_code);
      series=strip(_global_series);
      originator_name=strip(_originator_name);
      genetic_owner=strip(_genetic_owner_name);
      owner_group_code=input(_owner_group_code, best.);
      grouping_code=strip(_grouping_code__perennials_ddl_);
      parent_line_indicator=strip(_parent_line_indicator);
      short_code=strip(_short_code);
      variety=input(_foundation_variety_code, best.);
      variety_research_code=strip(_variety_research_code);
      breeder=strip(_breeder);
      lead_product_manager=strip(_lead_product_manager_res);

      /*GLOBAL*/
      GLOBAL_variety_name=strip(_GLOBAL_global_name);
      GLOBAL_current_plc=strip(_GLOBAL_global_plc);
      GLOBAL_future_plc=strip(_GLOBAL_GLOBAL_future_plc);
      GLOBAL_date_of_future_plc=input(_GLOBAL_GLOBAL_date_of_future_pl, date11.);
      GLOBAL_replace_by=input(_GLOBAL_GLOBAL_replace_by, best.);
      GLOBAL_replacement_date=input(_GLOBAL_GLOBAL_replacement_date_, date11.);
      GLOBAL_abc=strip(_global_GLOBAL_abc);
      GLOBAL_brand=strip(_GLOBAL_brand);

      /*EUROPE*/
      EUROPE_series_code=strip(_EUROPE_europe_series_code);
      EUROPE_series_description=strip(_EUROPE_europe_series_descriptio); /* use for SFE */
      EUROPE_variety_name=strip(_EUROPE_europe_name); /* use for SFE */
      EUROPE_current_plc=strip(_EUROPE_europe_plc); /* use for SFE */
      EUROPE_future_plc=strip(_EUROPE_europe_future_plc); /* use for SFE */
      EUROPE_date_of_future_plc=input(_EUROPE_europe_date_of_future_pl, date11.);
      EUROPE_replace_by=input(_EUROPE_europe_replace_by, best.);
      EUROPE_replacement_date=input(_EUROPE_europe_replacement_date_, date11.);
      EUROPE_abc=strip(_EUROPE_europe_abc);
      EUROPE_channel_eame=strip(_EUROPE_europe___channel_syt_eam);
      EUROPE_channel_floranova=strip(_EUROPE_europe___channel_florano);
      EUROPE_regional_product_manager=strip(_EUROPE_regional_product_manager);

      /*NAM*/
      NAM_series_code=strip(_NAM_n__america_series_code);
      NAM_series_description=strip(_NAM_n__america_series_descripti);
      NAM_variety_name=strip(_NAM_NAM_name);
      NAM_current_plc=strip(_NAM_NAM_plc);
      NAM_future_plc=strip(_NAM_NAM_future_plc);
      NAM_date_of_future_plc=input(_NAM_NAM_date_of_future_plc, date11.);
      NAM_replace_by=input(_NAM_NAM_replace_by, best.);
      NAM_replacement_date=input(_NAM_NAM_replacement_date, date11.);
      NAM_abc=strip(_NAM_NAM_abc);
      NAM_regional_product_manager=strip(_NAM_regional_product_manager_re);

      /*APAC*/
      APAC_series_code=strip(_APAC_apac_series_code);
      APAC_series_description=strip(_APAC_apac_series_description);  /* use for all reports about BI, JP, FN */
      APAC_variety_name=strip(_APAC_apac_name);  /* use for all reports about BI, JP, FN */
      APAC_current_plc=strip(_APAC_apac_plc);  /* use for all reports about BI, JP, FN */
      APAC_future_plc=strip(_APAC_apac_future_plc); /* use for all reports about BI, JP, FN */
      APAC_date_of_future_plc=input(_APAC_apac_date_of_future_plc, date11.);
      APAC_replace_by=input(_APAC_apac_replace_by, best.);
      APAC_replacement_date=input(_APAC_apac_replacement_date__d, date11.);
      APAC_abc=strip(_APAC_apac_abc);
      APAC_channel_bi=strip(_APAC_apac___channel_syt_bi);  /* use this for BI, JP, FN */
      APAC_channel_floranova=strip(_APAC_apac___channel_floranova);  /* do not use */
      APAC_channel_japan=strip(_APAC_apac___channel_japan);          /* do not use */
      APAC_regional_product_manager=strip(_APAC_regional_product_manager_r);

      /*AME*/
      AME_series_code=strip(_AME_ame_series_code);
      AME_series_description=strip(_AME_ame_series_description);
      AME_variety_name=strip(_AME_ame_name);
      AME_current_plc=strip(_AME_ame_plc);
      AME_future_plc=strip(_AME_ame_future_plc);
      AME_date_of_future_plc=input(_AME_ame_date_of_future_plc, date11.);
      AME_replace_by=input(_AME_ame_replace_by, best.);
      AME_replacement_date=input(_AME_ame_replacement_date__dd, date11.);
      AME_abc=strip(_AME_ame_abc);
      AME_channel_bi=strip(_AME_ame___channel_syt_bi);
      AME_channel_floranova=strip(_AME_ame___channel_floranova);
      AME_regional_product_manager=strip(_AME_regional_product_manager_re);

      /*LATAM*/
      LATAM_series_code=strip(_LATAM_latam_series_code);
      LATAM_series_description=strip(_LATAM_latam_series_description);
      LATAM_variety_name=strip(_LATAM_latam_name);
      LATAM_current_plc=strip(_LATAM_latam_plc);
      LATAM_future_plc=strip(_LATAM_latam_future_plc);
      LATAM_date_of_future_plc=input(_LATAM_latam_date_of_future_plc, date11.);
      LATAM_replace_by=input(_LATAM_latam_replace_by, best.);
      LATAM_replacement_date=input(_LATAM_latam_replacement_date__, date11.);
      LATAM_abc=strip(_LATAM_latam_abc);
      LATAM_channel_bi=strip(_LATAM_latam___channel_syt_bi);
      LATAM_channel_floranova=strip(_LATAM_latam___channel_floranov);
      LATAM_regional_product_manager=strip(_LATAM_regional_product_manager_);   
    
    run;

    %if &pmd_file.=1 %then %do;
      data DMIMPORT.PMD_ALL;
        set pmd1;   
      run;
    %end; %else %do;
      proc append base=DMIMPORT.PMD_ALL data=pmd1 force nowarn;
      run;
    %end;

  %end;

  /*temporary deduplication - should be solved in the source files*/
  proc sort data=DMIMPORT.PMD_ALL nodupkey;
    by variety;
  run;

  data DMIMPORT.PMD_GLOBAL(DROP=EUROPE: NAM: APAC: AME: LATAM:);
    length PMD_region $6.;
    set DMIMPORT.PMD_ALL(rename=(
                    GLOBAL_variety_name=Variety_name 
                    GLOBAL_current_plc=Current_PLC
                    GLOBAL_future_plc=Future_PLC
                    GLOBAL_date_of_future_plc=Future_PLC_Active_date
                    GLOBAL_replace_by=replace_by
                    GLOBAL_replacement_date=replacement_date
                    GLOBAL_abc=abc
                    GLOBAL_brand=brand
                    )); 
    PMD_region="GLOBAL";
    if ^missing(Variety_name) and Variety_name ^= "0" then output;
  run;

  data DMIMPORT.PMD_EUROPE(DROP=NAM: APAC: AME: LATAM:);
    length PMD_region $6.;
    set DMIMPORT.PMD_ALL(rename=(
                    EUROPE_variety_name=Variety_name 
                    EUROPE_current_plc=Current_PLC
                    EUROPE_future_plc=Future_PLC
                    EUROPE_date_of_future_plc=Future_PLC_Active_date
                    EUROPE_replace_by=replace_by
                    EUROPE_replacement_date=replacement_date
                    EUROPE_abc=abc
                    EUROPE_channel_eame=channel_EAME
                    EUROPE_channel_floranova=channel_FLORANOVA
                    EUROPE_regional_product_manager=regional_product_manager
                    )); 
    PMD_region="EUROPE";
    variety_name=coalescec(variety_name, GLOBAL_variety_name);
    if ^missing(Variety_name) and Variety_name ^= "0" then output;
  run;

  data DMIMPORT.PMD_NAM(DROP=EUROPE: APAC: AME: LATAM:);
    length PMD_region $6.;
    set DMIMPORT.PMD_ALL(rename=(
                    NAM_variety_name=Variety_name 
                    NAM_current_plc=Current_PLC
                    NAM_future_plc=Future_PLC
                    NAM_date_of_future_plc=Future_PLC_Active_date
                    NAM_replace_by=replace_by
                    NAM_replacement_date=replacement_date
                    NAM_abc=abc
                    NAM_regional_product_manager=regional_product_manager
                    )); 
    PMD_region="NAM";
    variety_name=coalescec(variety_name, GLOBAL_variety_name);
    if ^missing(Variety_name) and Variety_name ^= "0" then output;
  run;

  data DMIMPORT.PMD_APAC(DROP=EUROPE: NAM: AME: LATAM:);
    length PMD_region $6.;
    set DMIMPORT.PMD_ALL(rename=(
                    APAC_variety_name=Variety_name 
                    APAC_current_plc=Current_PLC
                    APAC_future_plc=Future_PLC
                    APAC_date_of_future_plc=Future_PLC_Active_date
                    APAC_replace_by=replace_by
                    APAC_replacement_date=replacement_date
                    APAC_abc=abc
                    APAC_channel_bi=channel_BI
                    APAC_channel_floranova=channel_FLORANOVA
                    APAC_channel_japan=channel_JAPAN
                    APAC_regional_product_manager=regional_product_manager
                    )); 
    PMD_region="APAC";
    variety_name=coalescec(variety_name, GLOBAL_variety_name);
    if ^missing(Variety_name) and Variety_name ^= "0" then output;
  run;

  data DMIMPORT.PMD_AME(DROP=EUROPE: NAM: APAC: LATAM:);
    length PMD_region $6.;
    set DMIMPORT.PMD_ALL(rename=(
                    AME_variety_name=Variety_name 
                    AME_current_plc=Current_PLC
                    AME_future_plc=Future_PLC
                    AME_date_of_future_plc=Future_PLC_Active_date
                    AME_replace_by=replace_by
                    AME_replacement_date=replacement_date
                    AME_abc=abc
                    AME_channel_bi=channel_BI
                    AME_channel_floranova=channel_FLORANOVA
                    AME_regional_product_manager=regional_product_manager
                    )); 
    PMD_region="AME";
    variety_name=coalescec(variety_name, GLOBAL_variety_name);
    if ^missing(Variety_name) and Variety_name ^= "0" then output;
  run;

  data DMIMPORT.PMD_LATAM(DROP=EUROPE: NAM: APAC: AME:);
    length PMD_region $6.;
    set DMIMPORT.PMD_ALL(rename=(
                    LATAM_variety_name=Variety_name 
                    LATAM_current_plc=Current_PLC
                    LATAM_future_plc=Future_PLC
                    LATAM_date_of_future_plc=Future_PLC_Active_date
                    LATAM_replace_by=replace_by
                    LATAM_replacement_date=replacement_date
                    LATAM_abc=abc
                    LATAM_channel_bi=channel_BI
                    LATAM_channel_floranova=channel_FLORANOVA
                    LATAM_regional_product_manager=regional_product_manager
                    )); 
    PMD_region="LATAM";
    variety_name=coalescec(variety_name, GLOBAL_variety_name);
    if ^missing(Variety_name) and Variety_name ^= "0" then output;
  run;

  data DMIMPORT.PMD_SFE(drop=channel:);
    length region $6.;
    set DMIMPORT.PMD_EUROPE;
    region='SFE';
    if channel_eame='Y' then output;
  run;

  /* bugfix - configuration adjustment */

/*    data DMIMPORT.PMD_FN(DROP=EUROPE: NAM: APAC: LATAM:);*/
/*    length region $6.;*/
/*    set DMIMPORT.PMD_ALL(rename=(*/
/*                    AME_variety_name=Variety_name */
/*                    AME_current_plc=Current_PLC*/
/*                    AME_future_plc=Future_PLC*/
/*                    AME_date_of_future_plc=Future_PLC_Active_date*/
/*                    EUROPE_replace_by=replace_by*/
/*                    EUROPE_replacement_date=replacement_date*/
/*                    EUROPE_abc=abc*/
/*                    AME_channel_bi=channel_BI*/
/*                    EUROPE_channel_floranova=channel_FLORANOVA*/
/*                    EUROPE_regional_product_manager=regional_product_manager*/
/*                    )); */
/*    region="FN";*/
/*    variety_name=coalescec(variety_name, GLOBAL_variety_name);*/
/*    if ^missing(Variety_name) and Variety_name ^= "0" and channel_floranova='Y' then output;*/
/*  run;*/

/*  data DMIMPORT.PMD_FN(drop=channel:);*/
/*    length region $6.;*/
/*    set DMIMPORT.PMD_EUROPE;*/
/*    region='FN';*/
/*    if channel_floranova='Y' then output;*/
/*  run;*/

  /* bugfix - configuration adjustment - BI, JP and FN use APAC configuration */
  
  data DMIMPORT.PMD_BI(drop=channel:);
    length region $6.;
    set DMIMPORT.PMD_APAC;
    region='BI';
	series_name_in_region=apac_series_description;  /* 23 JULY 2021 RMP feature request : regional series names */
	abc = APAC_abc;  								/* 23 JULY 2021 RMP feature request : regional ABC classes */
    if channel_bi='Y' then output;   
  run;

  data DMIMPORT.PMD_JP(drop=channel:);
    length region $6.;
    set DMIMPORT.PMD_APAC;
    region='JP';
	series_name_in_region=apac_series_description;  /* 23 JULY 2021 RMP feature request : regional series names */
	abc = APAC_abc;  								/* 23 JULY 2021 RMP feature request : regional ABC classes */
    if channel_japan='Y' then output; 
  run;


  data DMIMPORT.PMD_FN(drop=channel:);
    length region $6.;
    set DMIMPORT.PMD_APAC;
    region='FN';
	series_name_in_region=apac_series_description;  /* 23 JULY 2021 RMP feature request : regional series names */
	abc = APAC_abc;  								/* 23 JULY 2021 RMP feature request : regional ABC classes */
    if channel_floranova='Y' then output; 
  run;

%mend IMPORT_PMD_ASSORTMENT;

%macro PROCESS_PMD_ASSORTMENT();

  data PMD;
    length variety 8.;
    length Replace_by 8.;
    length Replacement_Date 8.;
    length genetics $11.;
    length Species_code $4.;

    set 
      DMIMPORT.PMD_SFE
      DMIMPORT.PMD_BI
      DMIMPORT.PMD_JP
      DMIMPORT.PMD_FN
/*      DMIMPORT.PMD_EUROPE*/
/*      DMIMPORT.PMD_NAM*/
/*      DMIMPORT.PMD_APAC*/
/*      DMIMPORT.PMD_AME*/
/*      DMIMPORT.PMD_LATAM*/
/*      DMIMPORT.PDM_GLOBAL*/
      ;    

    if genetic_owner in ('SYNGENTA VEG', 'SYNGENTA', 'SYNGENTAFL') then do;
      genetics='SYNGENTA';
    end; else do;
      genetics='Third party';
    end;
    if Parent_Line_Indicator='N' and 
       ^missing(variety) 
      then output;
  run;

  data PMD1;
    set PMD;
    format Future_PLC_Active_date Global_Future_PLC_Active_Date DDMMYYP10.;
    if upcase(strip(product_line)) ^in('PERENNIAL CUTTING' 'PERENNIAL SEED') then call missing(Grouping_Code);
    if strip(Grouping_Code) ^in('A', 'B', 'C', 'D', 'E', 'F', 'G') then call missing(Grouping_Code);
    hash_product_line=lowcase(strip(product_line));
    hash_species_name=lowcase(strip(Species));
    hash_Grouping_Code=lowcase(strip(Grouping_Code));
  run;

  /*apply seasons to assortment*/
  data PMD2 (drop=rc);
    set PMD1;
    length season_week_start season_week_end 8. product_line_group $20. plg_rule $10.;

    if _n_=1 then do;
      declare hash sg(dataset: 'dmimport.seasons_general');
        rc=sg.DefineKey ('hash_product_line');
        rc=sg.DefineData ('plg_rule', 'product_line_group', 'season_week_start', 'season_week_end');
        rc=sg.DefineDone();

      declare hash sgce(dataset: 'dmimport.seasons_Grouping_Code_exc');
        rc=sgce.DefineKey ('hash_product_line', 'hash_grouping_code');
        rc=sgce.DefineData ('plg_rule', 'product_line_group', 'season_week_start', 'season_week_end');
        rc=sgce.DefineDone();

      declare hash sse(dataset: 'dmimport.seasons_species_exc');
          rc=sse.DefineKey ('hash_product_line', 'hash_species_name');
          rc=sse.DefineData ('plg_rule', 'product_line_group', 'season_week_start', 'season_week_end');
          rc=sse.DefineDone();

      declare hash vct(dataset: 'dmimport.variety_class_table');
          rc=vct.DefineKey ('variety');
          rc=vct.DefineData ('species_code');
          rc=vct.DefineDone();
    end;

    rc=sg.find();   /*gets seasons from seasons_general*/
    rc=sgce.find(); /*gets seasons exceptions for product_group (overwrite general seasons for product line) from seasons_product_group_exc*/
    rc=sse.find();  /*gets seasons exceptions for species (overwrite general seasons for product line and exception for grouping code) from seasons_seasons_exc*/
    rc=vct.find();  /*lookup for the species_code in variety_class_table*/
  run;

  data DMPROC.PMD_assortment;
    set PMD2;
  run;


%mend PROCESS_PMD_ASSORTMENT;