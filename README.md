# googlesheet-read-write-by-key
Script for googlesheet. Reads from googlesheet1, which is then used to writes data on googlesheet2. It essentially combines two tables with the 'ASIN' serving as a foreign key.

## How it works
Reading: Script will read 'googlesheet1' file to map ASIN to appropriate data. Any row with blank ASIN cell will be ignored. If multiple ASINS are detected, the script will overwrite with the last ASIN found.

Writing: Script will not iterate through all sheets; only the active one. Script reads each row's ASIN, and fill out the data based on what was stored from the 'googlesheet1' file. It will write "undefined" if the 'googlesheet1' has ASIN, but doesn't have the other header. If it does have the other header, but the row has no val, then it will write a empty string.

![Alt text](/images/demo.jpg?raw=true)