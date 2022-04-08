/***********************************************************************/
/*Type: Process*/
/*Use: Used in import_all only (macro call  %PROCESS_MATERIAL_ASSORTMENT()*/
/*Purpose: Process FPS_ASSORTMENT and MATERIAL_CLASS_TABLE into 1 table*/
/*IN: dmimport.fps_assortment*/
/*    dmimport.material_class_table*/
/*    dmimport.seasons_species_exc*/
/*OUT: DMPROC.MATERIAL_ASSORTMENT*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";

%macro process_material_assortment();

  data material_class_table;
    length source_material_table $50.;
    set dmimport.material_class_table      
      (keep=Material_Desc
          Division
          Material
          Variety
          sub_unit
          region
          proc_stage
		  packing_type
        rename=(
          Material_Desc=bi_material_name
          Division=material_division
          proc_stage=process_stage
        ));
    source_material_table='material_class_table';
  run;

  data fps_assortment;
    length source_material_table $50.;
    set dmimport.fps_assortment
      (keep=
        Material_basic_description_EN
        Material_division
        PF_for_sales_text
        Sub_unit
        DM_Process_stage
        Curr_Mat_PLC
        Future_Material_PLC
        Material_Future_PLC_valid_from
        material
        Variety
        region
        replaced_by
        bulk_6a
      rename=(
        Material_basic_description_EN=fps_material_name
        PF_for_sales_text=product_form
        Curr_Mat_PLC=material_plc_current
        Future_Material_PLC=material_plc_future
        Material_Future_PLC_valid_from=material_plc_future_date
        DM_Process_stage=process_stage
      ));
    source_material_table='fps_assortment';
  run;

  data dmproc.material_assortment;
    retain
      source_material_table
      region
      variety
      material
      material_division
      sub_unit
      product_form
      process_stage
      fps_material_name
      bi_material_name;
      length process_stage $5.;
    set 
      material_class_table
      fps_assortment;
  run;

%mend process_material_assortment;