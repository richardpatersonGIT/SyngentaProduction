/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %IMPORT_PMD_ASSORTMENT(...)  %PROCESS_PMD_ASSORTMENT()*/
/*Purpose: Imports (IMPORT_PMD_ASSORTMENT) and process (PROCESS_PMD_ASSORTMENT) PMD Assortment Global files from folder (pmd_assortment_folder=) with .csv extension*/
/*IN: pmd_assortment(1 file), extension=csv, delimiter = ';'*/
/*    dmimport.seasons_general*/
/*    dmimport.seasons_Grouping_Code_exc*/
/*    dmimport.seasons_species_exc*/
/*OUT: DMPROC.PMD_assortment*/
/***********************************************************************/

%include "C:\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%macro IMPORT_PMD_ASSORTMENT(pmd_assortment_folder=);

  %read_folder(folder=%str(&pmd_assortment_folder.), filelist=pmd_assortment, ext_mask=csv);

  proc sql noprint;
  select path_file_name into :pmd_file trimmed from pmd_assortment where order=1;
  quit;

  data PMD_RAW;
    infile "&pmd_file." delimiter = ';' missover dsd firstobs=2 lrecl=32767;
    length 
      Entry $1.
      Species $29.
      Series $26.
      Color_Code $4.
      Color_Description $50.
      Research_Code $23.
      Zvarm $7.
      Variety_Number $8.
      Purpose $18.
      Originator_Name $50.
      Variety_Name_Global $53.
      Global_Current_PLC $2.
      Global_New_PLC $1.
      Global_Future_PLC $2.
      Global_Future_Name $31.
      Global_Future_PLC_Active_Date $11.
      Proposed_Global_Future_PLC $1.
      Proposed_Global_Future_Name $1.
      Proposed_Global_Ftr_PLC_Act_Dt $1.
      Variety_Name_EAME $60.
      EAME_Current_PLC $2.
      EAME_Future_PLC $2.
      EAME_Future_PLC_Active_date $11.
      EAME_Future_Name $36.
      Proposed_EAME_Future_PLC $1.
      Proposed_EAME_Future_Date $1.
      Proposed_EAME_Future_Name $1.
      Variety_Name_BI $46.
      Current_PLC_BI $2.
      Future_PLC_BI $2.
      Future_PLC_Active_date_BI $11.
      Future_Name_BI $18.
      Proposed_Future_PLC_BI $1.
      Proposed_Future_Date_BI $1.
      Proposed_Future_Name_BI $1.
      Variety_Name_NA $52.
      NA_Current_PLC $2.
      NA_Future_PLC $2.
      NA_Future_PLC_Active_Date $11.
      NA_Future_Name $10.
      Proposed_NA_Future_PLC $1.
      Proposed_NA_Future_Date $1.
      Proposed_NA_Future_Name $1.
      Variety_Name_Japan $46.
      Japan_Current_PLC $2.
      Japan_Future_PLC $2.
      Japan_Future_PLC_Active_Date $11.
      Japan_Future_Name $1.
      Proposed_Japan_Future_PLC $1.
      Proposed_Japan_Future_Date $1.
      Proposed_Japan_Future_Name $1.
      Genetic_Productivity_Trial $50.
      Remarks $100.
      _Replace_by $20.
      Replaces $20.
      _Replacement_Date_dd_mon_yyyy $11.
      Assortment_PDC_fill $1.
      Grouping_Code $4.
      Hardiness_heat_zone $6.
      Family $21.
      Product_Line $19.
      Registered_Name $50.
      Out_license_to $50.
      Breeder $34.
      Regional_Product_Manager_rsp $31.
      Lead_Product_Manager_rsp $19.
      Product_Forms $37.
      Free_Field2 $1.
      Free_Field3 $3.
      Awards $62.
      Royalty_Value_Euro_per_1000 $20.
      Royalty_Value_Dollar_per_1000 $5.
      PBR_PVP_PP_required_Region_year $50.
      Free_Field4 $1.
      Replacer1_variety_nr $8.
      Substitute1_Originator_name $54.
      Substitute1_Supplier $25.
      Substitute1_Global_Name $51.
      Substitute1_EAME_Name $56.
      Substitute1_NA_Name $40.
      Substitute1_Remark $100.
      Substitute1_Silent_replacer_YN $3.
      Replacer2_variety_nr $8.
      Substitute2_Originator_name $49.
      Substitute2_Supplier $16.
      Substitute2_Global_Name $36.
      Substitute2_EAME_Name $43.
      Substitute2_NA_Name $41.
      Substitute2_Remarks $100.
      Substitute2_Silent_replacer_YN $3.
      Replacer3_variety_nr $8.
      Substitute3_Originator_name $37.
      Substitute3_Supplier $16.
      Substitute3_Global_Name $27.
      Substitute3_EAME_Name $38.
      Substitute3_NA_Name $31.
      Substitute3_Remarks $56.
      Substitute3_Silent_replacer_YN $2.
      Genetic_Owner $10.
      Brand $6.
      Variety_Business_Short_Code $8.
      Owner_Group $2.
      Multiplication_Indicator $20.
      Parent_Line_Indicator $1.
      Parent_Female_Material_Number $8.
      Parent_Male_Material_Number $8.
      NAFTAISNewVariety $1.
      NAFTAColorCategory $1.
      NAFTAColorDescription $1.
      NAFTAHabit $1.
      NAFTAApplication $1.
      NAFTAGardenerApplication $1.
      NAFTAGardenSituation $1.
      NAFTASeason $1.
      NAFTALightExposure $1.
      NAFTAHeight $1.
      NAFTAWidth $1.
      NAFTADescription $1.
      NAFTAGardenerVarietyDescription $1.
      NAFTAFlowerCrop $21.
      NAFTAGardenerFlowerSeason $1.
      NAFTAWWWDisplay $5.
      Selection_Advice $90.
;
    input 
      Entry $
      Species $
      Series $
      Color_Code $
      Color_Description $
      Research_Code $
      Zvarm $
      Variety_Number $
      Purpose $
      Originator_Name $
      Variety_Name_Global $
      Global_Current_PLC $
      Global_New_PLC $
      Global_Future_PLC $
      Global_Future_Name $
      Global_Future_PLC_Active_Date $
      Proposed_Global_Future_PLC $
      Proposed_Global_Future_Name $
      Proposed_Global_Ftr_PLC_Act_Dt $
      Variety_Name_EAME $
      EAME_Current_PLC $
      EAME_Future_PLC $
      EAME_Future_PLC_Active_date $
      EAME_Future_Name $
      Proposed_EAME_Future_PLC $
      Proposed_EAME_Future_Date $
      Proposed_EAME_Future_Name $
      Variety_Name_BI $
      Current_PLC_BI $
      Future_PLC_BI $
      Future_PLC_Active_date_BI $
      Future_Name_BI $
      Proposed_Future_PLC_BI $
      Proposed_Future_Date_BI $
      Proposed_Future_Name_BI $
      Variety_Name_NA $
      NA_Current_PLC $
      NA_Future_PLC $
      NA_Future_PLC_Active_Date $
      NA_Future_Name $
      Proposed_NA_Future_PLC $
      Proposed_NA_Future_Date $
      Proposed_NA_Future_Name $
      Variety_Name_Japan $
      Japan_Current_PLC $
      Japan_Future_PLC $
      Japan_Future_PLC_Active_Date $
      Japan_Future_Name $
      Proposed_Japan_Future_PLC $
      Proposed_Japan_Future_Date $
      Proposed_Japan_Future_Name $
      Genetic_Productivity_Trial $
      Remarks $
      _Replace_by $
      Replaces $
      _Replacement_Date_dd_mon_yyyy $
      Assortment_PDC_fill $
      Grouping_Code $
      Hardiness_heat_zone $
      Family $
      Product_Line $
      Registered_Name $
      Out_license_to $
      Breeder $
      Regional_Product_Manager_rsp $
      Lead_Product_Manager_rsp $
      Product_Forms $
      Free_Field2 $
      Free_Field3 $
      Awards $
      Royalty_Value_Euro_per_1000 $
      Royalty_Value_Dollar_per_1000 $
      PBR_PVP_PP_required_Region_year $
      Free_Field4 $
      Replacer1_variety_nr $
      Substitute1_Originator_name $
      Substitute1_Supplier $
      Substitute1_Global_Name $
      Substitute1_EAME_Name $
      Substitute1_NA_Name $
      Substitute1_Remark $
      Substitute1_Silent_replacer_YN $
      Replacer2_variety_nr $
      Substitute2_Originator_name $
      Substitute2_Supplier $
      Substitute2_Global_Name $
      Substitute2_EAME_Name $
      Substitute2_NA_Name $
      Substitute2_Remarks $
      Substitute2_Silent_replacer_YN $
      Replacer3_variety_nr $
      Substitute3_Originator_name $
      Substitute3_Supplier $
      Substitute3_Global_Name $
      Substitute3_EAME_Name $
      Substitute3_NA_Name $
      Substitute3_Remarks $
      Substitute3_Silent_replacer_YN $
      Genetic_Owner $
      Brand $
      Variety_Business_Short_Code $
      Owner_Group $
      Multiplication_Indicator $
      Parent_Line_Indicator $
      Parent_Female_Material_Number $
      Parent_Male_Material_Number $
      NAFTAISNewVariety $
      NAFTAColorCategory $
      NAFTAColorDescription $
      NAFTAHabit $
      NAFTAApplication $
      NAFTAGardenerApplication $
      NAFTAGardenSituation $
      NAFTASeason $
      NAFTALightExposure $
      NAFTAHeight $
      NAFTAWidth $
      NAFTADescription $
      NAFTAGardenerVarietyDescription $
      NAFTAFlowerCrop $
      NAFTAGardenerFlowerSeason $
      NAFTAWWWDisplay $
      Selection_Advice $
      ;
  run;


  data DMIMPORT.PMD_ALL;
    set PMD_RAW (keep=
                    Species
                    Series
                    Variety_Number
                    Variety_Name_Global
                    Global_Current_PLC
                    Global_Future_PLC
                    Global_Future_Name
                    Global_Future_PLC_Active_Date
                    Variety_Name_EAME
                    EAME_Current_PLC
                    EAME_Future_PLC
                    EAME_Future_PLC_Active_date
                    EAME_Future_Name
                    Variety_Name_BI
                    Current_PLC_BI
                    Future_PLC_BI
                    Future_PLC_Active_date_BI
                    Future_Name_BI
                    Grouping_Code
                    Product_Line
                    Free_Field3
                    Brand
                    Parent_Line_Indicator
                    _Replace_by
                    _Replacement_Date_dd_mon_yyyy
                    Genetic_Owner
                    Variety_Name_Japan
                    Japan_Current_PLC
                    Japan_Future_PLC
                    Japan_Future_PLC_Active_Date
                    Japan_Future_Name
                    );    
  run;

  data DMIMPORT.PMD_FPS;
    set DMIMPORT.PMD_ALL(rename=(Variety_Name_EAME=Variety_name 
                    EAME_Current_PLC=Current_PLC
                    EAME_Future_PLC=Future_PLC
                    EAME_Future_PLC_Active_date=Future_PLC_Active_date
                    EAME_Future_Name=Future_Name
                    )); 
    length region $6.;
    region="FPS";
    call missing(Global_Current_PLC);
    call missing(Global_Future_PLC);
    call missing(Global_Future_PLC_Active_Date);
    if ^missing(Variety_name) and Variety_name ^= "0" then output;
  run;

  data DMIMPORT.PMD_BI;
    set DMIMPORT.PMD_ALL(rename=(Variety_Name_BI=Variety_name 
                    Current_PLC_BI=Current_PLC
                    Future_PLC_BI=Future_PLC
                    Future_PLC_Active_date_BI=Future_PLC_Active_date
                    Future_Name_BI=Future_Name

                    )); 
    length region $6.;
    region="BI";
    if ^missing(Variety_name) and Variety_name ^= "0" then output;
  run;
  
  data DMIMPORT.PMD_JP;
    set DMIMPORT.PMD_ALL(rename=(Variety_Name_Japan=Variety_name 
                    Japan_Current_PLC=Current_PLC
                    Japan_Future_PLC=Future_PLC
                    Japan_Future_PLC_Active_Date=Future_PLC_Active_date
                    Japan_Future_Name=Future_Name

                    )); 
    length region $6.;
    region="JP";
    if ^missing(Variety_name) and Variety_name ^= "0" then output;
  run;
  
  data DMIMPORT.PMD_WHS; /*no assortment yet for WHS, BI assortment taken as a filler for now*/
    set DMIMPORT.PMD_ALL(rename=(Variety_Name_BI=Variety_name 
                    Current_PLC_BI=Current_PLC
                    Future_PLC_BI=Future_PLC
                    Future_PLC_Active_date_BI=Future_PLC_Active_date
                    Future_Name_BI=Future_Name

                    )); 
    length region $6.;
    region="WHS";
    if ^missing(Variety_name) and Variety_name ^= "0" then output;
  run;
  

  data DMIMPORT.PMD_GLOBAL;
    set DMIMPORT.PMD_ALL(rename=(Variety_Name_Global=Variety_name 
                    Global_Future_Name=Future_Name
                    ));
    length region $6.; 
    region="GLOBAL";
    Current_PLC=Global_Current_PLC;
    Future_PLC=Global_Future_PLC;
    Future_PLC_Active_date=Global_Future_PLC_Active_Date;
    if ^missing(Variety_name) and Variety_name ^= "0" then output;
  run;

%mend IMPORT_PMD_ASSORTMENT;

%macro PROCESS_PMD_ASSORTMENT();

  data PMD (drop=Variety_Name_EAME
                EAME_Current_PLC
                EAME_Future_PLC
                EAME_Future_PLC_Active_date
                EAME_Future_Name
                Variety_Name_BI
                Current_PLC_BI
                Future_PLC_BI
                Future_PLC_Active_date_BI
                Future_Name_BI
                Variety_Name_Global
                Global_Future_Name
                variety_number
                );
    length variety 8.;
    length Replace_by 8.;
    length Replacement_Date 8.;
    length genetics $11.;
    format Replacement_Date yymmdd10.;
    length Species_code $4.;
    
    set DMIMPORT.PMD_FPS
        DMIMPORT.PMD_BI 
        DMIMPORT.PDM_GLOBAL
        DMIMPORT.PMD_JP
        DMIMPORT.PMD_WHS;
    variety=input(variety_number, 8.);
        
    if ^missing(_Replacement_Date_dd_mon_yyyy) then do;
      Replacement_Date=input(_Replacement_Date_dd_mon_yyyy, date11.);
    end;
    Replace_by=input(_Replace_by, 8.);
    if genetic_owner in ('SYNGENTA', 'SYNGENTA V', 'SYNGENTAFL') then do;
      genetics='SYNGENTA';
    end; else do;
      genetics='Third party';
    end;
    if Parent_Line_Indicator='N' then output;
  run;

  data PMD1(drop=_Future_PLC_Active_date _Global_Future_PLC_Active_Date);
    set PMD(rename=(Future_PLC_Active_date=_Future_PLC_Active_date Global_Future_PLC_Active_Date=_Global_Future_PLC_Active_Date));
    format Future_PLC_Active_date Global_Future_PLC_Active_Date DDMMYYP10.;
    if ^missing(_Future_PLC_Active_date) then Future_PLC_Active_date=input(_Future_PLC_Active_date, date11.);
    if ^missing(_Global_Future_PLC_Active_Date) then Global_Future_PLC_Active_Date=input(_Global_Future_PLC_Active_Date, date11.);
    if strip(product_line) ^in('PERENNIAL CUTTING' 'PERENNIAL SEED') then call missing(Grouping_Code);
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