*** Settings ***
Resource                ../resources/common.resource
Resource                ../resources/excel.resource
Test Teardown           Close All Excel Documents
Suite Setup             Setup Browser
Suite Teardown          End Suite

*** Test Cases ***
Read Excel Data 1
    [Documentation]     Read Salesfoce record data from excel from multiple sheets.
    ...                 ExcelLibrary keyword documentation:
    ...                 https://rawgit.com/peterservice-rnd/robotframework-excellib/master/docs/ExcelLibrary.html
    [Tags]              excel    data    salesforce
    Login
    SetConfig           DefaultTimeout    10

    ${file}=            Set Variable    ${CURDIR}/../files/salesforce_test_data.xlsx
    ${sheet_name}=      Set Variable    Accounts
    ${testname}=        Set Variable    Communications

    # Open existing workbook
    ${document}=        Open Excel Document    ${file}    excel
    Log To Console      ${document}

    ${header_row}=	    Read Excel Row	    row_num=1    sheet_name=${sheet_name}

    @{rows}=            Create List
    ${row_num}=         Set Variable    ${2}

    WHILE    True    limit=20
        ${row}=	            Read Excel Row	    row_num=${row_num}    sheet_name=${sheet_name}
        Log To Console      ${row}

        # Break from loop when row is empty, we have reached the end of the sheet
        IF    "${row}[0]" == "None"
            Log To Console      Empty row
            BREAK
        END

        ${row_num}=         Evaluate    ${row_num} + 1
        Log To Console      ${row_num}

        ${status}=         Run Keyword And Return Status
        ...                Should Contain Match    ${row}    ${testname}

        # Continue to next row if match not found
        IF    ${status} != ${True}
            CONTINUE
        END

        ${row_dict}=        Create Dictionary
        FOR    ${index}    ${header}    IN ENUMERATE    @{header_row}
            Set To Dictionary    ${row_dict}    ${header}    ${row}[${index}]
        END
        Append To List      ${rows}    ${row_dict}
    END

    Close All Excel Documents

    Log To Console      ${rows}
    Set Suite Variable  ${rows}

Read Excel Data 2
    [Tags]              excel    data    salesforce
    Login
    SetConfig           DefaultTimeout    10

    ${file}=            Set Variable    ${CURDIR}/../files/salesforce_test_data.xlsx
    ${sheet_name}=      Set Variable    Accounts
    ${testname}=        Set Variable    Communications

    @{excel_rows}=      Load Data Line    ${file}    ${sheet_name}    ${testname}

    Log To Console      ${excel_rows}
    Set Suite Variable  ${excel_rows}
