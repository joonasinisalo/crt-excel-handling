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

    RETURN              ${rows}
