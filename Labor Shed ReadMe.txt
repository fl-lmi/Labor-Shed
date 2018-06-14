The "Labor Shed Program.R" script generates a Microsoft Word document containing a near-complete Labor Shed Analysis for a single county in Florida. To run, the script requires
the flat LEHD Origin-Destination (main and auxiliary), Residence Area Characteristics, and Workforce Area Characteristics files, and a customized version of the most recent ACS
PEPANNRES table that contains population estimates for the nation, Florida, and all counties within Florida. The WID also needs to be added as a data source via 
ODBC Data Source Administrator to allow for the use of LAUS data. To do so, go to the Administrator and add a SQL Server driver. The name of the data source should be set to
'DEOSQLTEST' and the server should be listed as 'DEOSQLTEST02,53215'. 

Several variables in the script have to be initialized in order for the program to be executed successfully. Below is a list of these variables and explanations 
for how their values should be entered. To execute the script, place the cursor on the first line and press Ctrl + Alt + e. Once the script has been run and the Microsoft Word
document has been created, the following need to be manually entered into the document: a title, maps for the county, and the county's outflow ranking in the Executive Summary
section. The document can be found in the same folder as the script, with a file name in the format "*Area Name* Labor Shed.docx" 


Variables to enter: 

area - County FIPS code of selection area. FIPS codes should be entered as three-digit character strings within single quotation marks. For example, 
       if the County FIPS code is 001, it should be entered as '001'. 
    


area_name - Name of county. Name should be entered as a character string within single quotation marks. For example, if the county name is Alachua
            County, enter 'Alachua County'. 


html_area - This corresponds to the URL associated with the county QuickFacts table available through Census. Usually, the portion of the URL that
            references the selection area takes the form of the county name and the state name, all in lowercase and without spaces. For example,
            Alachua County would be 'alachuacountyflorida'. 


pop_data_range - Range of data to be displayed in the population table. Entered as year:year. For example, if data for 2010 - 2016 are being reported, enter 2010:2016.   


report_date - A character string for the date for which the analysis is being generated. For example, if the analysis is being done for the month
              of October in 2017, enter 'October 2017'.


wfr - Code for the Workforce Region where the county is found, entered as a six-digit character string. For example, Alachua County is located in 
      Workforce Region 9, so '000009' would be entered if a report for Alachua County were being done.


month - Two-digit character string representing the most recent month for which LAUS data exist. For example, if September is the latest
        month, enter '09'.   


year - An integer corresponding to the year for which the most recent LAUS data exist. 


month_name - Character string for the name of the month that was entered for the "month" variable. So if the month of interest is September, 
             enter 'September'. 

