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

    ${row_dict}=        Create Dictionary
    ${row_num}=         Set Variable    ${2}

    # COMMENT: Set the limit to what you think is the max amount of rows
    WHILE    True    limit=200
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

        FOR    ${index}    ${header}    IN ENUMERATE    @{header_row}
            Set To Dictionary    ${row_dict}    ${header}    ${row}[${index}]
        END
        BREAK
    END

    Close All Excel Documents

    FOR    ${key_value_tuple}    IN    &{row_dict}
    #Set every column name    "${${key_value_tuple}[0]}" as a separate Suite variable with the corresponding row value "${key_value_tuple}[1]"
        IF    '${key_value_tuple}[1]' != ''
            Set Suite Variable          ${${key_value_tuple}[0]}    ${key_value_tuple}[1]
        END
    END
