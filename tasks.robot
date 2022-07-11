*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.PDF


*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the intranet website
    login
    download the excel file
    Fill the form from data from excel
    collect the results
    export the table as pdf
    [Teardown]    logout and close Browser


*** Keywords ***
Open the intranet website
    Open Available Browser    https://robotsparebinindustries.com/

login
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

download the excel file
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=true

Fill the form for one person
    [Arguments]    ${sales_reps}
    Input Text    firstname    ${sales_reps}[First Name]
    Input Text    lastname    ${sales_reps}[Last Name]
    Input Text    salesresult    ${sales_reps}[Sales]
    Select From List By Value    salestarget    ${sales_reps}[Sales Target]
    Click Button    Submit

Fill the form from data from excel
    Open Workbook    SalesData.xlsx
    ${sales_reps}=    Read Worksheet As Table    header=true
    Close Workbook
    FOR    ${sales_reps}    IN    @{sales_reps}
        Fill the form for one person    ${sales_reps}
    END

collect the results
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png

export the table as pdf
    Wait Until Element Is Visible    id:sales-results
    ${sales_results_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales_results.pdf

 logout and close Browser
    Click Button    Log out
    Close Browser
