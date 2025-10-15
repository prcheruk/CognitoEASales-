
select a.install_at_party_id,a.quantity,a.item_name,a.inventory_item_id product_id,
a.serial_number,
a.ship_date,
a.install_location_id,
COVERED_STATUS coverage_status,
INSTANCE_STATUS_DESC,
   (select D_EXT_ATTR3
from apps.XXCCS_EGO_MTL_SY_ITEMS_EXT_B x
where x.attr_group_id=224
and x.organization_id=900000000
and x.inventory_item_id=a.inventory_item_id
and rownum=1) eol_date,
   (select D_EXT_ATTR10
from apps.XXCCS_EGO_MTL_SY_ITEMS_EXT_B x
where x.attr_group_id=224
and x.organization_id=900000000
and x.inventory_item_id=a.inventory_item_id
and rownum=1) eos_date,
(select mc.segment1 from
apps.mtl_categories_b mc,
       apps.mtl_item_categories mic,
       apps.mtl_category_sets mcs
       --XXSCM_PL_BU_PF_PID_MV
 where mc.category_id = mic.category_id
       AND mic.inventory_item_id = a.inventory_item_id
       AND mic.organization_id = 900000000
       AND mcs.category_set_id = mic.category_set_id
       AND mcs.category_set_name = 'PROD GROUP'
       ) PF,
b.cr_party_id party_id,b.cr_party_name party_name,cr_country_name party_country_name,
b.cav_id, b.cav_name, b.cav_bu_id,b.cav_bu_name
from APPS.XXCCS_DS_INSTANCE_DETAIL a, APPS.XXCCS_DS_SITE_GU_DENORM b
where  b.cav_id=&cav_id
and a.install_at_site_use_id=b.site_use_id
FETCH FIRST 1 ROWS ONLY
 