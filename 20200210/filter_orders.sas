/***********************************************************************/
/*Type: Utility*/
/*Use: Used as a macro*/
/*Purpose: Used to filter out the unwanted orders. */
/*IN: parameter in_table*/

/*OUT:*/
/*	parameter out_table*/
/*	table: orders_filtered_for_checking*/
/***********************************************************************/
%macro filter_orders(in_table=, out_table=);

  data &out_table. orders_filtered_for_checking;
    set &in_table.;
    if 
      (upcase(country)='JAPAN' and sls_org='JP03' and index(upcase(product_line_group), 'CU')>0) 
    then do;
      output orders_filtered_for_checking;
    end; else do;
      output &out_table.;
    end;
  run;

%mend filter_orders;