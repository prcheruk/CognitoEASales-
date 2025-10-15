select d.account_NAME, d.account_Number account_number, d.cust_account_id cust_act_id, g.party_id,e.party_name,
b.site_use_id site_use_id, decode(b.site_use_code,'SHIP_TO','SID','BID') site_type,
C.PARTY_SITE_ID, G.LOCATION_ID, L.aDDRESS1, l.ADDRESS2, L.ADDRESS3,L.CITY,L.POSTAL_CODE,L.STATE,L.COUNTRY,
cavh.*
from apps.hz_cust_site_uses_all b, -- uses of customer addresses
apps.HZ_CUST_ACCT_SITES_ALL c, -- customer addresses
apps.hz_cust_accounts_all d, -- customer accounts
apps.HZ_PARTY_SITES G,
apps.HZ_LOCATIONS L,
apps.hz_parties e, -- parties
apps.xxccs_Ds_site_gu_denorm cavh
where  cavh.cav_id=539790
and b.cust_acct_site_id = c.cust_acct_site_id
and d.cust_account_id=cavh.cust_account_id
--409462715
and c.cust_account_id = d.cust_account_id
and d.party_id = e.party_id
AND G.PARTY_SITE_ID=C.PARTY_SITE_ID
AND L.LOCATION_ID=G.LOCATiON_ID;

//Or 

select * from apps.xxccs_Ds_site_gu_denorm cavh
where  cavh.cav_id=539790