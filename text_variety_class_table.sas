/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %IMPORT_VARIETY_CLASS_TABLE(...)*/
/*Purpose: Imports Variety Class Table from folder (variety_class_table_folder=) with .txt extension*/
/*IN: Variety_class_table(multiple files), extension=txt, delimiter=tab*/
/*OUT: dmimport.variety_class_table*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro IMPORT_VARIETY_CLASS_TABLE(variety_class_table_folder=);

  %read_folder(folder=%str(&variety_class_table_folder.), filelist=variety_class_files, ext_mask=txt);

  proc datasets lib=work nolist;
    delete variety_class_table_all;
  run;

  proc sql noprint;
    select count(*) into :variety_class_files_cnt from variety_class_files;
  quit;

  %do mc_file=1 %to &variety_class_files_cnt.;

    proc sql noprint;
      select path_file_name into :variety_class_file trimmed from variety_class_files where order = &mc_file.;
    quit;

    data variety_class_table_raw;
      infile "&variety_class_file." dlm='09'x dsd MISSOVER lrecl=32767  firstobs=4;

      length 
        /*_Material $18.
        Material_Desc $40.
        Division $2.
        Material_Group $7.
        Cnvt_Unit $3.
        Product_Hierarchy $10.
        Basic_Material $1.
        Material_Type $4.
        Old_Material_Number $1.
        Product_Allocation $1.
        Variety_Number_New $1.
        Variety_Description $1.
        Variety_Number_Old $1.
        Variety_Division $1.
        Class $3.
        Class_Description $7.
        Batch_Class $3.
        Class_Type $1.
        Brand $1.
        Buom $5.
        Caliber $3.
        Generation $5.
        Germ_Level $1.
        Grouping $1.
        Label_Code $1.
        Material_Desc1 $1.
        Material_Nr $1.
        Packing_Type $1.
        Proc_Stage $1.
        _Quantity $1.
        Quom $1.
        Reference_Code $1.
        Species $1.
        Srg $1.
        State_Certif $1.
        Sub_Brand $3.
        Treatment $1.
        Tsw $1.
        Variance_Key $2.
        Variety_Name $1.
        _Variety_Nr $1.;*/

        _Material $18.
        Material_Desc $40.
        Division $2.
        Material_Group $7.
        Cnvt_Unit $3.
        Product_Hierarchy $10.
        Basic_Material $1.
        Material_Type $4.
        Old_Material_Number $1.
        Product_Allocation $1.
        Variety_Number_New $1.
        Variety_Description $1.
        Variety_Number_Old $1.
        Variety_Division $1.
        Class $3.
        Class_Description $7.
        Batch_Class $3.
        Class_Type $1.
        BARLEY_TYPE $1.
        Brand $1.
        Buom $5.
        BUSINESS_SHORTCODE $1.
        CERTIFICATION_CODE $1.
        CERTIFIED $1.
        CHANNEL_AGR_REQ $1.
        COUNTRY_OPEN $1.
        DESTINATION $1.
        ECOLOGY $1.
        FEM_NAT_RESISTANCE $1.
        GAST_MANUALLY_PLACE $1.
        GAST_PLANT_PATTERN $1.
        GAST_PREFERRED_GROWINGAREA1 $1.
        GAST_PREFERRED_GROWINGAREA2 $1.
        GAST_PREFERRED_GROWINGAREA3 $1.
        GAST_SPLIT_REQUIREMENTS $1.
        GENETIC_OWNER $1.
        GMO $1.
        HEAT_UNITS $1.
        HYBRID_OPINDIC $1.
        INPUT_TRAIT $1.
        LICENSOR $1.
        MAINTAINER_NR $1.
        Material_Desc1 $1.
        Material_Nr $1.
        MATURITY_DAYS $1.
        MATURITY_GROUP $1.
        MATURITY_SUBGROUP $1.
        MEGA_SEGMENT $1.
        MINIMUM_GERM $1.
        NAT_RESISTANCE $1.
        OLD_MATERIAL_NR $1.
        OLEIC_LVL $1.
        OUTPUT_TRAIT $1.
        PARENT1 $1.
        PARENT2 $1.
        PARENT3 $1.
        PARENT4 $1.
        PARENT5 $1.
        PARENT6 $1.
        PARENT_LINE $1.
        PARENT_NP_CODE $1.
        PARENT_PD_CODE $1.
        PATENT_NUMBER $1.
        PLC_GLOB $1.
        PLM_RATING $1.
        PRODUCT_CATEGORY $1.
        PROD_HIER $1.
        PROPRIETARY $1.
        PVP $1.
        RECIPROCAL_NR $1.
        REGISTRATION_NUMBER $1.
        REGISTRATION_STATUS $1.
        REGULATED_VARIETY $1.
        RELATIVE_MATURITY $1.
        RESEARCH_CODE $1.
        REVERSIBLE_PARENT $1.
        ROYALTY $1.
        SERIES_CODE_GLOBAL $1.
        SPECIES $4.
        STEW_AGR_REQ $1.
        SYNONYMS $1.
        TSW_CATEGORY $1.
        VARIETY_CATEGORY $1.
        VARIETY_CODE $1.
        X_CODE $1.;


    input 
    /*
      _Material $
      Material_Desc $
      Division $
      Material_Group $
      Cnvt_Unit $
      Product_Hierarchy $
      Basic_Material $
      Material_Type $
      Old_Material_Number $
      Product_Allocation $
      Variety_Number_New $
      Variety_Description $
      Variety_Number_Old $
      Variety_Division $
      Class $
      Class_Description $
      Batch_Class $
      Class_Type $
      Brand $
      Buom $
      Caliber $
      Generation $
      Germ_Level $
      Grouping $
      Label_Code $
      Material_Desc1 $
      Material_Nr $
      Packing_Type $
      Proc_Stage $
      _Quantity $
      Quom $
      Reference_Code $
      Species $
      Srg $
      State_Certif $
      Sub_Brand $
      Treatment $
      Tsw $
      Variance_Key $
      Variety_Name $
      _Variety_Nr $
    ;
    */

        _Material $
        Material_Desc $
        Division $
        Material_Group $
        Cnvt_Unit $
        Product_Hierarchy $
        Basic_Material $
        Material_Type $
        Old_Material_Number $
        Product_Allocation $
        Variety_Number_New $
        Variety_Description $
        Variety_Number_Old $
        Variety_Division $
        Class $
        Class_Description $
        Batch_Class $
        Class_Type $
        BARLEY_TYPE $
        Brand $
        Buom $
        BUSINESS_SHORTCODE $
        CERTIFICATION_CODE $
        CERTIFIED $
        CHANNEL_AGR_REQ $
        COUNTRY_OPEN $
        DESTINATION $
        ECOLOGY $
        FEM_NAT_RESISTANCE $
        GAST_MANUALLY_PLACE $
        GAST_PLANT_PATTERN $
        GAST_PREFERRED_GROWINGAREA1 $
        GAST_PREFERRED_GROWINGAREA2 $
        GAST_PREFERRED_GROWINGAREA3 $
        GAST_SPLIT_REQUIREMENTS $
        GENETIC_OWNER $
        GMO $
        HEAT_UNITS $
        HYBRID_OPINDIC $
        INPUT_TRAIT $
        LICENSOR $
        MAINTAINER_NR $
        Material_Desc1 $
        Material_Nr $
        MATURITY_DAYS $
        MATURITY_GROUP $
        MATURITY_SUBGROUP $
        MEGA_SEGMENT $
        MINIMUM_GERM $
        NAT_RESISTANCE $
        OLD_MATERIAL_NR $
        OLEIC_LVL $
        OUTPUT_TRAIT $
        PARENT1 $
        PARENT2 $
        PARENT3 $
        PARENT4 $
        PARENT5 $
        PARENT6 $
        PARENT_LINE $
        PARENT_NP_CODE $
        PARENT_PD_CODE $
        PATENT_NUMBER $
        PLC_GLOB $
        PLM_RATING $
        PRODUCT_CATEGORY $
        PROD_HIER $
        PROPRIETARY $
        PVP $
        RECIPROCAL_NR $
        REGISTRATION_NUMBER $
        REGISTRATION_STATUS $
        REGULATED_VARIETY $
        RELATIVE_MATURITY $
        RESEARCH_CODE $
        REVERSIBLE_PARENT $
        ROYALTY $
        SERIES_CODE_GLOBAL $
        SPECIES $
        STEW_AGR_REQ $
        SYNONYMS $
        TSW_CATEGORY $
        VARIETY_CATEGORY $
        VARIETY_CODE $
        X_CODE $;

    run;

    data variety_class_table (drop=_: rename=(Material_Desc=Variety_desc SPECIES=Species_code));
      set variety_class_table_raw;
      variety=input(substr(_material, 11, 8), 8.);
      region='FN';
      output;
      region='BI';
      output;
      region='JP';
      output;  
    run;

    proc append base=variety_class_table_all data=variety_class_table;
    run;

  %end;

  proc sort data=variety_class_table_all out=variety_class_table_sorted dupout=variety_class_table_dup nodupkey;
    by region variety;
  run;

  data dmimport.variety_class_table;
    set variety_class_table_sorted;
  run;

%mend IMPORT_VARIETY_CLASS_TABLE;