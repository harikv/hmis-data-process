# hmis-data-process
Downloads HMIS district level data files and processes them

# src/crawl.ts
Typescript file that uses puppeteer to download xls files from the data directory

# process-hmis-xls.ipynb
Python notebook that converts xls into xlsx and then extracts district level totals (without public/private and urban/rural breakdown) into a Pandas Dataframe and subsequently writes to Stata dta and csv files
