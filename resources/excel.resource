*** Settings ***
Library                 QForce
Library                 String
Library                 Collections
Library                 ExcelLibrary

*** Keywords ***
Load Data Line
    [Arguments]    ${file}    ${sheet_name}    ${testname}

    # Open existing workbook
    ${document}=        Open Excel Document    ${file}    excel
