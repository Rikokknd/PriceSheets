# PriceSheets

This is a real business solution. It takes a price list of goods(in .xlsx file format) from client-provided location, and uploads it to Google Sheets file, splitting the sheet to multiple sheets by keyword, formatting and applying styles in the process.

<!-- 
Cron settings:
*/15 8-20 * * *
/var/www/u0853380/data/priceSheets/bin/python /var/www/u0853380/data/priceSheets/pricelist.py >/dev/null 2>&1
 -->