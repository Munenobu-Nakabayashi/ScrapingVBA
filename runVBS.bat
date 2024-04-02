@echo off

setlocal enabledelayedexpansion

pushd C:\scraping

cscript .\runVBS.vbs "C:\scraping\scraping_IPA.xlsm" "sheet1.main" 2>> .\error.log
rem cscript .\enterKey.vbs
cscript .\runVBS.vbs "C:\scraping\scraping_JPA.xlsm" "sheet1.main" 2>> .\error.log
rem cscript .\enterKey.vbs
cscript .\runVBS.vbs "C:\scraping\scraping_JVN.xlsm" "sheet1.main" 2>> .\error.log
rem cscript .\enterKey.vbs
cscript .\runVBS.vbs "C:\scraping\scraping_JVN_iPedia.xlsm" "sheet1.main" 2>> .\error.log
rem cscript .\enterKey.vbs
cscript .\runVBS.vbs "C:\scraping\scraping_JVNDB.xlsm" "sheet1.main" 2>> .\error.log
rem cscript .\enterKey.vbs
cscript .\runVBS.vbs "C:\scraping\scraping_Microsoft.xlsm" "sheet1.main" 2>> .\error.log
rem cscript .\enterKey.vbs
cscript .\runVBS.vbs "C:\scraping\scraping_NISC.xlsm" "sheet1.main" 2>> .\error.log
rem cscript .\enterKey.vbs
cscript .\runVBS.vbs "C:\scraping\scraping_NIST_CVE.xlsm" "sheet1.main" 2>> .\error.log
rem cscript .\enterKey.vbs

popd

exit 0