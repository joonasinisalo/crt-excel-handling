*** Settings ***
Resource                ../resources/common.resource
Resource                ../resources/excel.resource
Test Teardown           Close All Excel Documents
Suite Setup             Setup Browser
Suite Teardown          End Suite

*** Test Cases ***
Read Excel Data
    [Documentation]     Read Salesfoce record data from excel from multiple sheets.
    ...                 ExcelLibrary keyword documentation:
    ...                 https://rawgit.com/peterservice-rnd/robotframework-excellib/master/docs/ExcelLibrary.html
    [Tags]              excel    data    salesforce
    ${file}=            Set Variable    ${CURDIR}/../data/salesforce_test_data.xlsx
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

        # Break from loop when row is empty
        IF    "${row}[0]" == "None"
            Log To Console      Empty row
            BREAK
        END

        ${row_num}=         Evaluate    ${row_num} + 1
        Log To Console      ${row_num}

        ${row_dict}=        Create Dictionary
        FOR    ${index}    ${header}    IN ENUMERATE    @{header_row}
            Set To Dictionary    ${row_dict}    ${header}    ${row}[${index}]
        END
        Append To List      ${rows}    ${row_dict}
    END

    Log To Console      ${rows}
    Set Suite Variable  ${rows}

    Close All Excel Documents
