/***********************************************************************/
/*Type: Import*/
/*Use: Used in import_all only (macro call %IMPORT_ORDERS(...) and PROCESS_ORDERS()*/
/*Purpose: Imports (IMPORT_ORDERS) and process (PROCESS_ORDERS) Sales Orders files from folder (orders_folder=) with .txt extension*/
/*IN: sales orders(multiple files), extension=txt, delimiter = ';'*/
/*    dmimport.FPS_assortment*/
/*    dmimport.Country_lookup*/
/*    dmimport.Material_class_table*/
/*    dmproc.PMD_assortment*/
/*OUT: dmiproc.orders_all*/
/*     dmimport.orders_rejected*/
/*     dmimport.order_duplicates*/
/***********************************************************************/

%include "C:\SAS\APPLICATIONS\SAS\configuration.sas";
%include "&sas_applications_folder.\read_folder.sas";

%MACRO order_import_variant_17(filename=);

  data orders_file_raw;
    infile "&filename." delimiter = ';' MISSOVER DSD lrecl=32767 firstobs=2;

    length
      Order_type  $4.
      Sls_doc_nr  $8.
      Sls_org  $4.
      DC  $2.
      Sls_div  $2.
      Sls_off  $4.
      Sls_grp  $3.
      Shp_plnt  $4.
      Shp_Pnt  $4.
      Route  $6.
      POnr  $20.
      Soldto_nr  $8.
      Soldto_name  $35.
      Cust_segm  $2.
      Shipto_nr  $8.
      Shipto_name  $35.
      Shipto_cntry  $2.
      Depot_Nr  $8.
      Depot_Name  $35.
      _Salrep_nr  $25.
      _Salrep_name  $28.
      Mat_div  $2.
      _Line_nr  $6.
      _Matnr  $8.
      Matdescr  $40.
      Mat_Sls_Txt  $72.
      Mat_grp  $9.
      _Line_req_deldte  $10.
      _SchedLine_Cnf_deldte  $10.
      _Hdr_req_deldte  $10.
      _Line_crdte  $10.
      _Doc_dte  $10.
      Ord_rsn  $3.
      _Order_Week  $7.
      _Ord_qty  $15.
      _Cnf_qty  $15.
      Qty_uom  $2.
      Itm_cnf_sts  $1.
      _Var_nr  $8.
      Var_descr  $40.
      Rsn_rej_cd  $2.
      Del_prio  $2.
      Sls_ord_Del_blck  $2.
      Sls_ord_bill_blck  $2.
      Itm_overall_sts  $1.
      Itm_del_sts  $1.
      Sls_ord_cred_sts  $1.
      Itm_usr_sts_cd  $4.
      _PR00  $15.
      PR00_Curr  $3.
      _Itm_Net_val  $15.
      Net_value_curr  $3.
      ABC_class  $2.
      Process_Order_Number  $12.
      _Scheduled_Release_Date  $10.
      _Scheduled_Start_Date  $10.
      _Scheduled_Finish_Date  $10.
      _Confirmed_Release_Date  $10.
      _Confirmed_Start_Date  $10.
      _Confirmed_Finish_Date  $10.
      InBatch_Number  $10.
      OutBatch_Number  $10.
      Supply_Status  $1.
      Prc_Ord_System_Status  $50.
      Prc_Ord_User_Stat  $4.
      filename  $40.
      ;

  input
    Order_type $
    Sls_doc_nr $
    Sls_org $
    DC $
    Sls_div $
    Sls_off $
    Sls_grp $
    Shp_plnt $
    Shp_Pnt $
    Soldto_nr $
    Soldto_name $
    Cust_segm $
    Shipto_nr $
    Shipto_name $
    Shipto_cntry $
    Mat_div $
    _Line_nr $
    _Matnr $
    Matdescr $
    Mat_Sls_Txt $
    Mat_grp $
    _Line_req_deldte $
    _SchedLine_Cnf_deldte $
    _Hdr_req_deldte $
    _Order_Week $
    _Line_crdte $
    _Doc_dte $
    Ord_rsn $
    _Ord_qty $
    _Cnf_qty $
    Qty_uom $
    _Var_nr $
    Var_descr $
    Rsn_rej_cd $
    Del_prio $
    Sls_ord_Del_blck $
    Sls_ord_bill_blck $
    Itm_cnf_sts $
    Itm_overall_sts $
    Itm_del_sts $
    Sls_ord_cred_sts $
    Itm_usr_sts_cd $
    _PR00 $
    PR00_Curr $
    _Itm_Net_val $
    Net_value_curr $
    ABC_class $
    Route $
    POnr $
    Depot_Nr $
    Depot_Name $
    _Salrep_nr $
    _Salrep_name $
;
    filename=strip("&filename.");
  Run;

  data _null_;
    sleep=sleep(0);
  run;

  proc append base=orders_raw_all data=orders_file_raw force;
  run;

%MEND order_import_variant_17;

%MACRO order_import_variant_18(filename=);

  data orders_file_raw;
    infile "&filename." delimiter = ';' MISSOVER DSD lrecl=32767 firstobs=2;

    length
      Order_type  $4.
      Sls_doc_nr  $8.
      Sls_org  $4.
      DC  $2.
      Sls_div  $2.
      Sls_off  $4.
      Sls_grp  $3.
      Shp_plnt  $4.
      Shp_Pnt  $4.
      Route  $6.
      POnr  $20.
      Soldto_nr  $8.
      Soldto_name  $35.
      Cust_segm  $2.
      Shipto_nr  $8.
      Shipto_name  $35.
      Shipto_cntry  $2.
      Depot_Nr  $8.
      Depot_Name  $35.
      _Salrep_nr  $25.
      _Salrep_name  $28.
      Mat_div  $2.
      _Line_nr  $6.
      _Matnr  $8.
      Matdescr  $40.
      Mat_Sls_Txt  $72.
      Mat_grp  $9.
      _Line_req_deldte  $10.
      _SchedLine_Cnf_deldte  $10.
      _Hdr_req_deldte  $10.
      _Line_crdte  $10.
      _Doc_dte  $10.
      Ord_rsn  $3.
      _Order_Week  $7.
      _Ord_qty  $15.
      _Cnf_qty  $15.
      Qty_uom  $2.
      Itm_cnf_sts  $1.
      _Var_nr  $8.
      Var_descr  $40.
      Rsn_rej_cd  $2.
      Del_prio  $2.
      Sls_ord_Del_blck  $2.
      Sls_ord_bill_blck  $2.
      Itm_overall_sts  $1.
      Itm_del_sts  $1.
      Sls_ord_cred_sts  $1.
      Itm_usr_sts_cd  $4.
      _PR00  $15.
      PR00_Curr  $3.
      _Itm_Net_val  $15.
      Net_value_curr  $3.
      ABC_class  $2.
      Process_Order_Number  $12.
      _Scheduled_Release_Date  $10.
      _Scheduled_Start_Date  $10.
      _Scheduled_Finish_Date  $10.
      _Confirmed_Release_Date  $10.
      _Confirmed_Start_Date  $10.
      _Confirmed_Finish_Date  $10.
      InBatch_Number  $10.
      OutBatch_Number  $10.
      Supply_Status  $1.
      Prc_Ord_System_Status  $50.
      Prc_Ord_User_Stat  $4.
      filename  $40.
      ;

  input
    Order_type $ 
    Sls_doc_nr $ 
    Sls_org $ 
    DC $ 
    Sls_div $ 
    Sls_off $ 
    Sls_grp $ 
    Shp_plnt $ 
    Shp_Pnt $ 
    Route $ 
    POnr $ 
    Soldto_nr $ 
    Soldto_name $ 
    Cust_segm $ 
    Shipto_nr $ 
    Shipto_name $ 
    Shipto_cntry $ 
    Depot_Nr $ 
    Depot_Name $ 
    _Salrep_nr $ 
    _Salrep_name $ 
    Mat_div $ 
    _Line_nr $ 
    _Matnr $ 
    Matdescr $ 
    Mat_Sls_Txt $ 
    Mat_grp $ 
    _Line_req_deldte $ 
    _SchedLine_Cnf_deldte $ 
    _Hdr_req_deldte $ 
    _Line_crdte $ 
    _Doc_dte $ 
    Ord_rsn $ 
    _Order_Week $ 
    _Ord_qty $ 
    _Cnf_qty $ 
    Qty_uom $ 
    Itm_cnf_sts $ 
    _Var_nr $ 
    Var_descr $ 
    Rsn_rej_cd $ 
    Del_prio $ 
    Sls_ord_Del_blck $ 
    Sls_ord_bill_blck $ 
    Itm_overall_sts $ 
    Itm_del_sts $ 
    Sls_ord_cred_sts $ 
    Itm_usr_sts_cd $ 
    _PR00 $ 
    PR00_Curr $ 
    _Itm_Net_val $ 
    Net_value_curr $ 
    ABC_class $ 
    Process_Order_Number $ 
    _Scheduled_Release_Date $ 
    _Scheduled_Start_Date $ 
    _Scheduled_Finish_Date $ 
    _Confirmed_Release_Date $ 
    _Confirmed_Start_Date $ 
    _Confirmed_Finish_Date $ 
    InBatch_Number $ 
    OutBatch_Number $ 
    Supply_Status $ 
    Prc_Ord_System_Status $ 
    Prc_Ord_User_Stat $;
    filename=strip("&filename.");
  Run;

  data _null_;
    sleep=sleep(0);
  run;

  proc append base=orders_raw_all data=orders_file_raw force;
  run;

%MEND order_import_variant_18;

%MACRO order_import_variant_201925(filename=);

  data orders_file_raw(drop=__:);
    infile "&filename." delimiter = ';' MISSOVER DSD lrecl=32767 firstobs=2;

    length
      Order_type  $4.
      Sls_doc_nr  $8.
      Sls_org  $4.
      DC  $2.
      Sls_div  $2.
      Sls_off  $4.
      Sls_grp  $3.
      Shp_plnt  $4.
      Shp_Pnt  $4.
      Route  $6.
      POnr  $20.
      Soldto_nr  $8.
      Soldto_name  $35.
      Cust_segm  $2.
      Shipto_nr  $8.
      Shipto_name  $35.
      Shipto_cntry  $2.
      Depot_Nr  $8.
      Depot_Name  $35.
      _Salrep_nr  $25.
      _Salrep_name  $28.
      Mat_div  $2.
      _Line_nr  $6.
      _Matnr  $8.
      Matdescr  $40.
      Mat_Sls_Txt  $72.
      Mat_grp  $9.
      _Line_req_deldte  $10.
      _SchedLine_Cnf_deldte  $10.
      _Hdr_req_deldte  $10.
      _Line_crdte  $10.
      _Doc_dte  $10.
      Ord_rsn  $3.
      _Order_Week  $7.
      _Ord_qty  $15.
      _Cnf_qty  $15.
      Qty_uom  $2.
      Itm_cnf_sts  $1.
      _Var_nr  $8.
      Var_descr  $40.
      Rsn_rej_cd  $2.
      Del_prio  $2.
      Sls_ord_Del_blck  $2.
      Sls_ord_bill_blck  $2.
      Itm_overall_sts  $1.
      Itm_del_sts  $1.
      Sls_ord_cred_sts  $1.
      Itm_usr_sts_cd  $4.
      _PR00  $15.
      PR00_Curr  $3.
      _Itm_Net_val  $15.
      Net_value_curr  $3.
      ABC_class  $2.
      Process_Order_Number  $12.
      _Scheduled_Release_Date  $10.
      _Scheduled_Start_Date  $10.
      _Scheduled_Finish_Date  $10.
      _Confirmed_Release_Date  $10.
      _Confirmed_Start_Date  $10.
      _Confirmed_Finish_Date  $10.
      InBatch_Number  $10.
      OutBatch_Number  $10.
      Supply_Status  $1.
      Prc_Ord_System_Status  $50.
      Prc_Ord_User_Stat  $4.
      __schedule_line_blckd_for_dlvry $1.
      __freeze_flag $1.
      __binding_delivery_date $1.
      __sales_order_item_text $1.
      __vendor_mat_no $1.
      filename  $40.
      ;

  input
    Order_type $ 
    Sls_doc_nr $ 
    Sls_org $ 
    DC $ 
    Sls_div $ 
    Sls_off $ 
    Sls_grp $ 
    Shp_plnt $ 
    Shp_Pnt $ 
    Route $ 
    POnr $ 
    Soldto_nr $ 
    Soldto_name $ 
    Cust_segm $ 
    Shipto_nr $ 
    Shipto_name $ 
    Shipto_cntry $ 
    Depot_Nr $ 
    Depot_Name $ 
    _Salrep_nr $ 
    _Salrep_name $ 
    Mat_div $ 
    _Line_nr $ 
    _Matnr $ 
    Matdescr $ 
    Mat_Sls_Txt $ 
    Mat_grp $ 
    _Line_req_deldte $ 
    _SchedLine_Cnf_deldte $ 
    _Hdr_req_deldte $ 
    _Line_crdte $ 
    _Doc_dte $ 
    Ord_rsn $ 
    _Order_Week $ 
    _Ord_qty $ 
    _Cnf_qty $ 
    Qty_uom $ 
    Itm_cnf_sts $ 
    _Var_nr $ 
    Var_descr $ 
    Rsn_rej_cd $ 
    Del_prio $ 
    Sls_ord_Del_blck $ 
    Sls_ord_bill_blck $ 
    Itm_overall_sts $ 
    Itm_del_sts $ 
    Sls_ord_cred_sts $ 
    Itm_usr_sts_cd $ 
    _PR00 $ 
    PR00_Curr $ 
    _Itm_Net_val $ 
    Net_value_curr $ 
    ABC_class $ 
    Process_Order_Number $ 
    _Scheduled_Release_Date $ 
    _Scheduled_Start_Date $ 
    _Scheduled_Finish_Date $ 
    _Confirmed_Release_Date $ 
    _Confirmed_Start_Date $ 
    _Confirmed_Finish_Date $ 
    InBatch_Number $ 
    OutBatch_Number $ 
    Supply_Status $ 
    Prc_Ord_System_Status $ 
    Prc_Ord_User_Stat $
    __schedule_line_blckd_for_dlvry $
    __freeze_flag $
    __binding_delivery_date $
    __sales_order_item_text $
    __vendor_mat_no $;
    filename=strip("&filename.");
  Run;

  data _null_;
    sleep=sleep(0);
  run;

  proc append base=orders_raw_all data=orders_file_raw force;
  run;

%MEND order_import_variant_201925;

%MACRO order_import_variant_201926(filename=);

  data orders_file_raw(drop=__:);
    infile "&filename." delimiter = ';' MISSOVER DSD lrecl=32767 firstobs=2;

    length
      Order_type  $4.
      Mat_div  $2.
      Shp_plnt  $4.
      Shp_Pnt  $4.
      POnr  $20.
      _Line_crdte  $10.
      Soldto_nr  $8.
      Soldto_name  $35.
      Shipto_nr  $8.
      Shipto_name  $35.
      Sls_doc_nr  $8.
      _Line_nr  $6.
      _Matnr  $8.
      Matdescr  $40.
      _Ord_qty  $15.
      _Cnf_qty  $15.
      Qty_uom  $2.
      _SchedLine_Cnf_deldte  $10.
      _Line_req_deldte  $10.
      Itm_del_sts  $1.
      _PR00  $15.
      _Itm_Net_val  $15.
      Sls_ord_Del_blck  $2.
      Sls_ord_bill_blck  $2.
      Sls_org  $4.
      DC  $2.
      Sls_div  $2.
      Sls_off  $4.
      Sls_grp  $3.
      Route  $6.
      Cust_segm  $2.
      Shipto_cntry  $2.
      Mat_Sls_Txt  $72.
      Mat_grp  $9.
      _Hdr_req_deldte  $10.
      _Doc_dte  $10.
      Ord_rsn  $3.
      Itm_cnf_sts  $1.
      _Var_nr  $8.
      Var_descr  $40.
      Rsn_rej_cd  $2.
      Del_prio  $2.
      Itm_overall_sts  $1.
      Sls_ord_cred_sts  $1.
      Itm_usr_sts_cd  $4.
      PR00_Curr  $3.
      Net_value_curr  $3.
      Depot_Nr  $8.
      Depot_Name  $35.
      _Salrep_nr  $25.
      _Salrep_name  $28.
      _Order_Week  $7.
      ABC_class  $2.
      Process_Order_Number  $12.
      _Scheduled_Release_Date  $10.
      _Scheduled_Start_Date  $10.
      _Scheduled_Finish_Date  $10.
      _Confirmed_Release_Date  $10.
      _Confirmed_Start_Date  $10.
      _Confirmed_Finish_Date  $10.
      InBatch_Number  $10.
      OutBatch_Number  $10.
      Supply_Status  $1.
      Prc_Ord_System_Status  $50.
      Prc_Ord_User_Stat  $4.
      __schedule_line_blckd_for_dlvry $1.
      __freeze_flag $1.
      __binding_delivery_date $1.
      __sales_order_item_text $1.
      filename  $40.
      ;

  input
    Order_type  $
    Mat_div  $
    Shp_plnt  $
    Shp_Pnt  $
    POnr  $
    _Line_crdte  $
    Soldto_nr  $
    Soldto_name  $
    Shipto_nr  $
    Shipto_name  $
    Sls_doc_nr  $
    _Line_nr  $
    _Matnr  $
    Matdescr  $
    _Ord_qty  $
    _Cnf_qty  $
    Qty_uom  $
    _SchedLine_Cnf_deldte  $
    _Line_req_deldte  $
    Itm_del_sts  $
    _PR00  $
    _Itm_Net_val  $
    Sls_ord_Del_blck  $
    Sls_ord_bill_blck  $
    Sls_org  $
    DC  $
    Sls_div  $
    Sls_off  $
    Sls_grp  $
    Route  $
    Cust_segm  $
    Shipto_cntry  $
    Mat_Sls_Txt  $
    Mat_grp  $
    _Hdr_req_deldte  $
    _Doc_dte  $
    Ord_rsn  $
    Itm_cnf_sts  $
    _Var_nr  $
    Var_descr  $
    Rsn_rej_cd  $
    Del_prio  $
    Itm_overall_sts  $
    Sls_ord_cred_sts  $
    Itm_usr_sts_cd  $
    PR00_Curr  $
    Net_value_curr  $
    Depot_Nr  $
    Depot_Name  $
    _Salrep_nr  $
    _Salrep_name  $
    _Order_Week  $
    ABC_class  $
    Process_Order_Number  $
    _Scheduled_Release_Date  $
    _Scheduled_Start_Date  $
    _Scheduled_Finish_Date  $
    _Confirmed_Release_Date  $
    _Confirmed_Start_Date  $
    _Confirmed_Finish_Date  $
    InBatch_Number  $
    OutBatch_Number  $
    Supply_Status  $
    Prc_Ord_System_Status  $
    Prc_Ord_User_Stat  $
    __schedule_line_blckd_for_dlvry $
    __freeze_flag $
    __binding_delivery_date $
    __sales_order_item_text $;
    filename=strip("&filename.");
  Run;

  data _null_;
    sleep=sleep(0);
  run;

  proc append base=orders_raw_all data=orders_file_raw force;
  run;

%MEND order_import_variant_201926;

%MACRO order_import_variant_202010(filename=);

  data orders_file_raw;
    infile "&filename." delimiter = ';' MISSOVER DSD lrecl=32767 firstobs=2;

    length
      Order_type  $4.
      Sls_doc_nr  $8.
      Sls_org  $4.
      DC  $2.
      Sls_div  $2.
      Sls_off  $4.
      Sls_grp  $3.
      Shp_plnt  $4.
      Shp_Pnt  $4.
      Route  $6.
      POnr  $20.
      Soldto_nr  $8.
      Soldto_name  $35.
      Cust_segm  $2.
      Shipto_nr  $8.
      Shipto_name  $35.
      Shipto_cntry  $2.
      Depot_Nr  $8.
      Depot_Name  $35.
      _Salrep_nr  $25.
      _Salrep_name  $28.
      Mat_div  $2.
      _Line_nr  $6.
      _Matnr  $8.
      Matdescr  $40.
      Mat_Sls_Txt  $72.
      Mat_grp  $9.
      _Line_req_deldte  $10.
      _SchedLine_Cnf_deldte  $10.
      _Hdr_req_deldte  $10.
      _Line_crdte  $10.
      _Doc_dte  $10.
      Ord_rsn  $3.
      _Order_Week  $7.
      _Ord_qty  $15.
      _Cnf_qty  $15.
      Qty_uom  $2.
      Itm_cnf_sts  $1.
      _Var_nr  $8.
      Var_descr  $40.
      Rsn_rej_cd  $2.
      Del_prio  $2.
      Sls_ord_Del_blck  $2.
      Sls_ord_bill_blck  $2.
      Itm_overall_sts  $1.
      Itm_del_sts  $1.
      Sls_ord_cred_sts  $1.
      Itm_usr_sts_cd  $4.
      _PR00  $15.
      PR00_Curr  $3.
      _Itm_Net_val  $15.
      Net_value_curr  $3.
      ABC_class  $2.
      Process_Order_Number  $12.
      _Scheduled_Release_Date  $10.
      _Scheduled_Start_Date  $10.
      _Scheduled_Finish_Date  $10.
      _Confirmed_Release_Date  $10.
      _Confirmed_Start_Date  $10.
      _Confirmed_Finish_Date  $10.
      InBatch_Number  $10.
      OutBatch_Number  $10.
      Supply_Status  $1.
      Prc_Ord_System_Status  $50.
      Prc_Ord_User_Stat  $4.
      Vendor_Mat_No $1.
      Shortcode $1.
      filename  $40.
      ;

  input
    Order_type $ 
    Sls_doc_nr $ 
    Sls_org $ 
    DC $ 
    Sls_div $ 
    Sls_off $ 
    Sls_grp $ 
    Shp_plnt $ 
    Shp_Pnt $ 
    Route $ 
    POnr $ 
    Soldto_nr $ 
    Soldto_name $ 
    Cust_segm $ 
    Shipto_nr $ 
    Shipto_name $ 
    Shipto_cntry $ 
    Depot_Nr $ 
    Depot_Name $ 
    _Salrep_nr $ 
    _Salrep_name $ 
    Mat_div $ 
    _Line_nr $ 
    _Matnr $ 
    Matdescr $ 
    Mat_Sls_Txt $ 
    Mat_grp $ 
    _Line_req_deldte $ 
    _SchedLine_Cnf_deldte $ 
    _Hdr_req_deldte $ 
    _Line_crdte $ 
    _Doc_dte $ 
    Ord_rsn $ 
    _Order_Week $ 
    _Ord_qty $ 
    _Cnf_qty $ 
    Qty_uom $ 
    Itm_cnf_sts $ 
    _Var_nr $ 
    Var_descr $ 
    Rsn_rej_cd $ 
    Del_prio $ 
    Sls_ord_Del_blck $ 
    Sls_ord_bill_blck $ 
    Itm_overall_sts $ 
    Itm_del_sts $ 
    Sls_ord_cred_sts $ 
    Itm_usr_sts_cd $ 
    _PR00 $ 
    PR00_Curr $ 
    _Itm_Net_val $ 
    Net_value_curr $ 
    ABC_class $ 
    Process_Order_Number $ 
    _Scheduled_Release_Date $ 
    _Scheduled_Start_Date $ 
    _Scheduled_Finish_Date $ 
    _Confirmed_Release_Date $ 
    _Confirmed_Start_Date $ 
    _Confirmed_Finish_Date $ 
    InBatch_Number $ 
    OutBatch_Number $ 
    Supply_Status $ 
    Prc_Ord_System_Status $ 
    Prc_Ord_User_Stat $
    Vendor_Mat_No $
    Shortcode $;
    filename=strip("&filename.");
  Run;

  data _null_;
    sleep=sleep(0);
  run;

  proc append base=orders_raw_all data=orders_file_raw force;
  run;

%MEND order_import_variant_202010;

%MACRO unknown_order_file_structure();

  %put "ERROR: Not recognized orders file structure from SAP, file &orders_file_name. will not be imported";

%MEND unknown_order_file_structure;

%macro PROCESS_ORDERS();

  data orders(drop=reject) dmimport.orders_rejected;
    set dmimport.orders_all;
    length reject 8.;
    /*reject*/
    if sls_div='24' then reject=1;
    if mat_grp='YMAGOODS' or mat_grp='YMASERVIC' then reject=1;
    if qty_uom='G' then reject=1;
    if sls_doc_nr='1164001' and matdescr='PEPZ ACAPULCO CASCADE' then reject=1; /*1 line of incorrect data*/
    if Soldto_name='SAMPLE CUSTOMER' then reject=1; /*sample customer*/
    if soldto_nr='10029017' and shipto_nr='10036720' and substr(matdescr, 1, 4)='MANZ' then reject=1; /*EXCEPTION - Delete all rows from Soldto nr.: 10029017 (Syngenta Seeds SAS Promo Fleurs) in combined with Ship to nr. 10036720 (Cultius Tolra, S.L.)*/

    if missing(reject) then output orders;
    if reject=1 then output dmimport.orders_rejected;
  run;

  data orders_filter;
    set orders;
    if mat_div in ('6B', '6C') and order_type in ('ZYPD', 'ZFD1', 'ZYPL', 'ZMTO') then output;
    if mat_div='6A' and order_type in ('YQOR', 'ZMTO') then output;
  run;

  data orders_proper (drop=_: 
                      rename=
                        (  matnr=material 
                          var_nr=variety));
    set orders_filter(rename=(_order_week=order_week_org));
    format Line_req_deldte SchedLine_Cnf_deldte Hdr_req_deldte Line_crdte Doc_dte Scheduled_Release_Date Scheduled_Start_Date Scheduled_Finish_Date Confirmed_Release_Date Confirmed_Start_Date Confirmed_Finish_Date yymmdd10.;
    length Matnr Var_nr Ord_qty Cnf_qty salrep_nr line_nr 8.;
    format PR00 Itm_Net_val 15.2;
    length unique_code $10.;

    /*dates*/
    if _Line_req_deldte^="00.00.0000" then Line_req_deldte=input(_Line_req_deldte,ddmmyy10.);
    if _SchedLine_Cnf_deldte^="00.00.0000" then SchedLine_Cnf_deldte=input(_SchedLine_Cnf_deldte,ddmmyy10.);
    if _Hdr_req_deldte^="00.00.0000" then Hdr_req_deldte=input(_Hdr_req_deldte,ddmmyy10.);
    if _Line_crdte^="00.00.0000" then Line_crdte=input(_Line_crdte,ddmmyy10.);
    if _Doc_dte^="00.00.0000" then Doc_dte=input(_Doc_dte,ddmmyy10.);
    if _Scheduled_Release_Date^="00.00.0000" then Scheduled_Release_Date=input(_Scheduled_Release_Date,ddmmyy10.);
    if _Scheduled_Start_Date^="00.00.0000" then Scheduled_Start_Date=input(_Scheduled_Start_Date,ddmmyy10.);
    if _Scheduled_Finish_Date^="00.00.0000" then Scheduled_Finish_Date=input(_Scheduled_Finish_Date,ddmmyy10.);
    if _Confirmed_Release_Date^="00.00.0000" then Confirmed_Release_Date=input(_Confirmed_Release_Date,ddmmyy10.);
    if _Confirmed_Start_Date^="00.00.0000" then Confirmed_Start_Date=input(_Confirmed_Start_Date,ddmmyy10.);
    if _Confirmed_Finish_Date^="00.00.0000" then Confirmed_Finish_Date=input(_Confirmed_Finish_Date,ddmmyy10.);

    /*numeric*/
    matnr=input(_Matnr, 8.);
    var_nr=input(_var_nr, 8.);
    line_nr=input(_line_nr, 8.);

    if index(_Ord_qty,"-")>1 then Ord_qty=-input(scan(_Ord_qty,1,'-') ,comma15.2);  
      else Ord_qty=input(_Ord_qty,comma15.2);
    if index(_Cnf_qty,"-")>1 then Cnf_qty=-input(scan(_Cnf_qty,1,'-') ,comma15.2); 
      else Cnf_qty=input(_Cnf_qty,comma15.2);
    if index(_PR00,"-")>1 then PR00=-input(scan(_PR00,1,'-') ,comma15.2);
      else PR00=input(_PR00,comma15.2);
    if index(_Itm_Net_val,"-")>1 then Itm_Net_val=-input(scan(_Itm_Net_val,1,'-') ,comma15.2);
      else Itm_Net_val=input(_Itm_Net_val,comma15.2);

    /*mixed up salrep_nr and salrep_name in source data*/
    if notdigit(strip(_salrep_nr))>0 then do;
      salrep_nr=input(_salrep_name, 8.);
      salrep_name=_salrep_nr;
    end; else do;
      salrep_nr=_salrep_nr;
      salrep_name=_salrep_name;
    end;

    unique_code=cats(sls_org, sls_off, shipto_cntry);

  run;

  data orders_regions(drop=rc);
    set orders_proper;
    length region $3. territory $3. country $6.;

    if _n_=1 then do;
      declare hash cl(dataset: 'dmimport.Country_lookup');
        rc=cl.DefineKey ('unique_code');
        rc=cl.DefineData ('region', 'territory', 'country');
        rc=cl.DefineDone();
      declare hash stl(dataset: 'dmimport.Soldto_nr_lookup');
        rc=stl.DefineKey ('soldto_nr');
        rc=stl.DefineData ('region', 'territory', 'country');
        rc=stl.DefineDone();
    end;

    rc=cl.find(); /*gets territory and country from country_lookup*/
    if region^='SFE' then do;
      rc=stl.find();/*gets territory and country from soldto_nr_lookup (if found overwrite the country_lookup)*/
    end;

  run;

  data orders_sub_unit(drop=rc);
    set orders_regions;
    length sub_unit 8. PF_for_sales_text $25. process_stage $5.;

    if _n_=1 then do;
      declare hash material_assortment(dataset: 'dmproc.material_assortment(rename=(product_form=PF_for_sales_text))');
        rc=material_assortment.DefineKey ('region', 'material');
        rc=material_assortment.DefineData ('sub_unit', 'process_stage', 'pf_for_sales_text');
        rc=material_assortment.DefineDone();     
    end;
    
    rc=material_assortment.find(); 

/*    if ^missing(sub_unit) then do;
      historical_sales=cnf_qty * sub_unit;
      if rsn_rej_cd in ('23', '60', '78', '64') then do;
        actual_sales=ord_qty * sub_unit;
      end; else do;
        actual_sales=historical_sales;
      end;
    end;
*/
    
    if ^missing(sub_unit) then do;
      if region='SFE' and rsn_rej_cd in ('ZR', '01', '02', '23', '57', '60', '64', '78') then do;
        historical_sales=0;
        actual_sales=ord_qty * sub_unit;
      end; 
     else if region in ('BI' 'JP' 'FN') and rsn_rej_cd in ('ZR', '01', '02', '23', '57', '68', '78') then do;
        historical_sales=0;
        actual_sales=ord_qty * sub_unit;
     end;
     else do;
        historical_sales=cnf_qty * sub_unit;
       actual_sales=historical_sales;
      end;
    end;

    br_actual_sales=0;
    br_historical_sales=0;

    if ^missing(sub_unit) then do;
      br_actual_sales=ord_qty * sub_unit;
      br_historical_sales=cnf_qty * sub_unit;
    end;

  run;

  data dmproc.orders_seasons(drop=rc);
    set orders_sub_unit;
    length  season_week_start season_week_end 
             Order_season_start order_year order_season order_month_season order_week Order_yweek order_month 8.;
    length product_line_group $20. Product_Line $19. plg_rule $10. species $29. series $26. Species_code $4.;
    format Order_season_start yymmdd10.;
    if _n_=1 then do;
      declare hash pmd_assortment(dataset: 'dmproc.PMD_assortment');
        rc=pmd_assortment.DefineKey ('region', 'variety');
        rc=pmd_assortment.DefineData ('season_week_start', 'season_week_end', 'product_line_group', 'product_line', 'plg_rule', 'species', 'series', 'species_code');
        rc=pmd_assortment.DefineDone();
    end;

    rc=pmd_assortment.find(); 

    
    order_year=year(SchedLine_Cnf_deldte);
    order_month=month(SchedLine_Cnf_deldte);
    order_yweek=input(substr(put(SchedLine_Cnf_deldte, weekv9.), 1, 4), 4.);
    order_week=input(substr(put(SchedLine_Cnf_deldte, weekv9.), 6, 2), 2.);
    if ^missing(season_week_start) then do;
      order_season_start=input(put(order_year, 4.)||'W'||put(season_week_start, z2.)||'01', weekv9.);
      SchedLine_Cnf_deldte_ym=input(put(year(SchedLine_Cnf_deldte),4.)||put(month(SchedLine_Cnf_deldte),z2.), 6.);
      order_season_start_ym=input(put(year(order_season_start),4.)||put(month(order_season_start),z2.), 6.);
      if SchedLine_Cnf_deldte >= order_season_start then do;
        order_season=order_year; 
      end; else do;
        order_season=order_year-1;
      end;

      if order_season_start_ym < SchedLine_Cnf_deldte_ym or (order_season_start_ym=SchedLine_Cnf_deldte_ym and day(order_season_start)<=15) then do;
        order_month_season=order_year;
      end; else do;
        order_month_season=order_year-1;
      end;
    end;

    /*if order_week=53 then order_week=52; - there is no week 53 in data anymore*/ 
  run;

  data orders_all;
    set dmproc.orders_seasons;
    length delivery_week delivery_season delivery_year delivery_month delivery_month_season 8.;
    delivery_week=put(order_yweek, 4.)||put(order_week, z2.);
    delivery_month=put(order_year, 4.)||put(order_month, z2.);
    delivery_year=order_year;
    delivery_season=order_season;
    delivery_month_season=order_month_season;
  run;

  /*varieties excluded - volmary - FRANCE */
  data dmproc.orders_all;
    set orders_all;
    if variety not in (
      70094938,
      70094940,
      70094941,
      70094942,
      70094943,
      70094944,
      70094949,
      70094950,
      70094951,
      70094954,
      70094955,
      70094957,
      70094958,
      70094960,
      70094961,
      70094962,
      70094963,
      70094964,
      70094965,
      70094966,
      70094967,
      70094968,
      70094969,
      70094971,
      70094972,
      70094973,
      70094974,
      70094976,
      70094977,
      70094980,
      70094981,
      70094983,
      70094985,
      70094987,
      70094988,
      70094989,
      70094990,
      70094991,
      70094996,
      70095001,
      70095002,
      70095004,
      70095005,
      70095006,
      70095007,
      70095008,
      70095009,
      70095011,
      70095012,
      70095015,
      70095017,
      70095018,
      70095019,
      70095021,
      70095022,
      70095024,
      70095025,
      70095026,
      70095027,
      70095028,
      70095029,
      70095030,
      70095031,
      70095032,
      70095033,
      70095034,
      70095035,
      70095037,
      70095038,
      70095040,
      70095041,
      70095043,
      70095044,
      70095046,
      70095047,
      70095048,
      70095049,
      70095050,
      70095051,
      70095052,
      70095053,
      70095054,
      70095055,
      70095056,
      70095057,
      70095058,
      70095059,
      70095060,
      70095061,
      70095064,
      70095066,
      70095067,
      70095068,
      70095069,
      70095070,
      70095071,
      70095073,
      70095074,
      70095075,
      70095077,
      70095078,
      70095079,
      70095080,
      70095081,
      70095082,
      70095084,
      70095085,
      70095086,
      70095087,
      70095089,
      70095090,
      70095091,
      70095092,
      70095093,
      70095094,
      70095095,
      70095096,
      70095097,
      70095098,
      70095099,
      70095100,
      70095101,
      70095102,
      70095103,
      70095104,
      70095105,
      70095106,
      70095108,
      70095109,
      70095110,
      70095111,
      70095112,
      70095113,
      70095114,
      70095115,
      70095116,
      70095117,
      70095118,
      70095119,
      70095120,
      70095121,
      70095122,
      70095123,
      70095124,
      70095125,
      70095126,
      70095127,
      70095131,
      70095132,
      70095133,
      70095134,
      70095135,
      70095136,
      70095137,
      70095138,
      70095139,
      70095140,
      70095141,
      70095142,
      70095144,
      70095145,
      70095147,
      70095149,
      70095151,
      70095153,
      70095154,
      70095155,
      70095156,
      70095161,
      70095163,
      70095165,
      70095166,
      70095167,
      70095168,
      70095169,
      70095170,
      70095171,
      70095172,
      70095173,
      70095174,
      70095175,
      70095176,
      70095177,
      70095178,
      70095179,
      70095180,
      70095181,
      70095182,
      70095183,
      70095184,
      70095185,
      70095186,
      70095187,
      70095188,
      70095189,
      70095190,
      70095191,
      70095192,
      70095194,
      70095195,
      70095197,
      70095198,
      70095199,
      70095200,
      70095201,
      70095202,
      70095203,
      70095204,
      70095205,
      70095206,
      70095207,
      70095208,
      70095209,
      70095210,
      70095211,
      70095212,
      70095213,
      70095214,
      70095215,
      70095217,
      70095218,
      70095219,
      70095220,
      70095221,
      70095222,
      70095223,
      70095224,
      70095227,
      70095228,
      70095229,
      70095230,
      70095231,
      70095232,
      70095233,
      70095234,
      70095235,
      70095236,
      70095237,
      70095238,
      70095241,
      70095243,
      70095244,
      70095245,
      70095283,
      70095284,
      70095286,
      70095287,
      70095288,
      70095289,
      70095290,
      70095291,
      70095293,
      70095294,
      70095295,
      70095296,
      70095297,
      70095298,
      70095299,
      70095300,
      70095302,
      70095303,
      70095304,
      70095306,
      70095309,
      70095310,
      70095311,
      70095312,
      70095313,
      70095314,
      70095315,
      70095316,
      70095317,
      70095318,
      70095319,
      70095320,
      70095321,
      70095322,
      70095323,
      70095324,
      70095325,
      70095327,
      70095328,
      70095329,
      70095330,
      70095331,
      70095332,
      70095333,
      70095334,
      70095335,
      70095336,
      70095337,
      70095344,
      70095345,
      70095859,
      70095889,
      70096342,
      70101810,
      70101841,
      70101842
      ) then output;
  run;

%mend PROCESS_ORDERS;

%macro IMPORT_ORDERS(orders_folder=);

  proc datasets lib=work nolist;
    delete orders_raw_all orders_raw;
  run;

  %read_folder(folder=%str(&orders_folder.), filelist=orders_list, ext_mask=txt);

  proc sql noprint;
    select count(*) into :order_files_cnt from orders_list;
  quit;

  %do order_file=1 %to &order_files_cnt.;

    proc sql noprint;
      select path_file_name into :orders_file_name trimmed from orders_list where order = &order_file.;
    quit;

    data column_check;
      length orders_macro $32. columns variant_17_columns variant_18_columns $1000. ;
      infile "&orders_file_name." MISSOVER DSD obs=1 lrecl=32767;
      input columns $;
      variant_17_columns='Order type;Sls doc nr;Sls org;DC;Sls div;Sls off;Sls grp;Shp plnt;Shp Pnt;Soldto nr;Soldto name;Cust segm;Shipto nr;Shipto name;Shipto cntry;Mat div;Line nr;Matnr;Matdescr;Mat Sls Txt;Mat grp;Line req deldte;SchedLine Cnf deldte;Hdr req deldte;Order Week;Line crdte;Doc dte;Ord rsn;Ord qty;Cnf qty;Qty uom;Var nr;Var descr;Rsn rej cd;Del prio;Sls ord Del blck;Sls ord bill blck;Itm cnf sts;Itm overall sts;Itm del sts;Sls ord cred sts;Itm usr sts cd;PR00;PR00 Curr;Itm Net val;Net value curr;ABC class;Route;POnr;Depot Nr;Depot Name;Salrep nr;Salrep name';
      variant_18_columns='Order type;Sls doc nr;Sls org;DC;Sls div;Sls off;Sls grp;Shp plnt;Shp Pnt;Route;POnr;Soldto nr;Soldto name;Cust segm;Shipto nr;Shipto name;Shipto cntry;Depot Nr;Depot Name;Salrep nr;Salrep name;Mat div;Line nr;Matnr;Matdescr;Mat Sls Txt;Mat grp;Line req deldte;SchedLine Cnf deldte;Hdr req deldte;Line crdte;Doc dte;Ord rsn;Order Week;Ord qty;Cnf qty;Qty uom;Itm cnf sts;Var nr;Var descr;Rsn rej cd;Del prio;Sls ord Del blck;Sls ord bill blck;Itm overall sts;Itm del sts;Sls ord cred sts;Itm usr sts cd;PR00;PR00 Curr;Itm Net val;Net value curr;ABC class;Process Order Number;Scheduled Release Date;Scheduled Start Date;Scheduled Finish Date;Confirmed Release Date;Confirmed Start Date;Confirmed Finish Date;InBatch Number;OutBatch Number;Supply Status;Prc. Ord. System Status;Prc Ord User Stat.;Schedule line blocked for delivery;Freeze Flag;Binding Delivery Date;Sales order Item text';
      variant_201925_columns='Order type;Sls doc nr;Sls org;DC;Sls div;Sls off;Sls grp;Shp plnt;Shp Pnt;Route;POnr;Soldto nr;Soldto name;Cust segm;Shipto nr;Shipto name;Shipto cntry;Depot Nr;Depot Name;Salrep nr;Salrep name;Mat div;Line nr;Matnr;Matdescr;Mat Sls Txt;Mat grp;Line req deldte;SchedLine Cnf deldte;Hdr req deldte;Line crdte;Doc dte;Ord rsn;Order Week;Ord qty;Cnf qty;Qty uom;Itm cnf sts;Var nr;Var descr;Rsn rej cd;Del prio;Sls ord Del blck;Sls ord bill blck;Itm overall sts;Itm del sts;Sls ord cred sts;Itm usr sts cd;PR00;PR00 Curr;Itm Net val;Net value curr;ABC class;Process Order Number;Scheduled Release Date;Scheduled Start Date;Scheduled Finish Date;Confirmed Release Date;Confirmed Start Date;Confirmed Finish Date;InBatch Number;OutBatch Number;Supply Status;Prc. Ord. System Status;Prc Ord User Stat.;Schedule line blocked for delivery;Freeze Flag;Binding Delivery Date;Sales order Item text;Vendor Mat. No.';
      variant_201926_columns='Order type;Mat div;Shp plnt;Shp Pnt;POnr;Line crdte;Soldto nr;Soldto name;Shipto nr;Shipto name;Sls doc nr;Line nr;Matnr;Matdescr;Ord qty;Cnf qty;Qty uom;SchedLine Cnf deldte;Line req deldte;Itm del sts;PR00;Itm Net val;Sls ord Del blck;Sls ord bill blck;Sls org;DC;Sls div;Sls off;Sls grp;Route;Cust segm;Shipto cntry;Mat Sls Txt;Mat grp;Hdr req deldte;Doc dte;Ord rsn;Itm cnf sts;Var nr;Var descr;Rsn rej cd;Del prio;Itm overall sts;Sls ord cred sts;Itm usr sts cd;PR00 Curr;Net value curr;Depot Nr;Depot Name;Salrep nr;Salrep name;Order Week;ABC class;Process Order Number;Scheduled Release Date;Scheduled Start Date;Scheduled Finish Date;Confirmed Release Date;Confirmed Start Date;Confirmed Finish Date;InBatch Number;OutBatch Number;Supply Status;Prc. Ord. System Status;Prc Ord User Stat.;Schedule line blocked for delivery;Freeze Flag;Binding Delivery Date;Sales order Item text';
      variant_202010_columns='Order type;Sls doc nr;Sls org;DC;Sls div;Sls off;Sls grp;Shp plnt;Shp Pnt;Route;POnr;Soldto nr;Soldto name;Cust segm;Shipto nr;Shipto name;Shipto cntry;Depot Nr;Depot Name;Salrep nr;Salrep name;Mat div;Line nr;Matnr;Matdescr;Mat Sls Txt;Mat grp;Line req deldte;SchedLine Cnf deldte;Hdr req deldte;Line crdte;Doc dte;Ord rsn;Order Week;Ord qty;Cnf qty;Qty uom;Itm cnf sts;Var nr;Var descr;Rsn rej cd;Del prio;Sls ord Del blck;Sls ord bill blck;Itm overall sts;Itm del sts;Sls ord cred sts;Itm usr sts cd;PR00;PR00 Curr;Itm Net val;Net value curr;ABC class;Process Order Number;Scheduled Release Date;Scheduled Start Date;Scheduled Finish Date;Confirmed Release Date;Confirmed Start Date;Confirmed Finish Date;InBatch Number;OutBatch Number;Supply Status;Prc. Ord. System Status;Prc Ord User Stat.;Schedule line blocked for delivery;Freeze Flag;Binding Delivery Date;Sales order Item text;Vendor Mat. No.;Shortcode';
      orders_macro='unknown_order_file_structure';
      if strip(variant_17_columns)=strip(columns) then orders_macro='order_import_variant_17';
      if strip(variant_18_columns)=strip(columns) then orders_macro='order_import_variant_18';
      if strip(variant_201925_columns)=strip(columns) then orders_macro='order_import_variant_201925';
      if strip(variant_201926_columns)=strip(columns) then orders_macro='order_import_variant_201926';
      if strip(variant_202010_columns)=strip(columns) then orders_macro='order_import_variant_202010';
    run;

    proc sql noprint;
      select orders_macro into :orders_macro from column_check;
    quit;

    %&orders_macro.(filename=%str(&orders_file_name.));

  %end;

  proc sort data=orders_raw_all out=orders_raw_all_sorted dupout=dmimport.order_duplicates nodupkey;
  by _all_;
  run;

  data dmimport.orders_all;
    set orders_raw_all_sorted;
  run;

%mend IMPORT_ORDERS;