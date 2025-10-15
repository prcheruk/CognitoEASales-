select  gu_name,d.cav_id,cav_name, cav_bu_name, party_name party_count, ib_covered_product_list,
ib_uncovered_product_list,ib_never_covered_product_list
 from apps.xxccs_Ds_site_gu_denorm d 
 where gu_name = &gu_name
 and cav_id is not null
 