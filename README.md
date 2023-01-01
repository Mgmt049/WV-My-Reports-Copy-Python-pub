# WV-My-Reports-Copy-Python-pub
 public python automation to distribute report files for a certain "on-prem" application
 
#############################################################################################
#Constants to populate before execution:
#############################################################################################
## LDAP_search_base: populate with LDAP search base - ou: organizational unit, o: organization, c: country, dc: domain
##Example: "OU=<organizational unit/container>,DC=<domain>,DC=com"
##AD_domain: populate as "<domain>.com"
##logon_domainL populate as "<domain>\\"
#############################################################################################
LDAP_search_base = ""
AD_domain = ""
logon_domain = ""
#############################################################################################