*** Settings ***
Library                 QForce
Library                 String
Library                 Collections
Library                 ExcelLibrary

# ExcelLibrary keyword documentation:
# https://rawgit.com/peterservice-rnd/robotframework-excellib/master/docs/ExcelLibrary.html


*** Keywords ***
Read Excel Row to Variables
    [Documentation]     Finds a row from an excel that matches the identifier and creates suite variables
    ...                 using column headers as variable names and cell values as variable values.
    [Arguments]         ${file_arg}    ${sheet_name_arg}    ${row_identifier_arg}

    # Open existing workbook
    ${document}=        Open Excel Document    ${file_arg}    excel

    ${header_row}=	    Read Excel Row	    row_num=1    sheet_name=${sheet_name_arg}

    &{row_dict}=        Create Dictionary
    ${row_num}=         Set Variable    ${2}

    # Set the limit to what you think is the max amount of rows
    WHILE    True    limit=200
        ${row}=	            Read Excel Row	    row_num=${row_num}    sheet_name=${sheet_name_arg}

        # Break from loop when row is empty, we have reached the end of the sheet
        IF    "${row}[0]" == "None"
            Log To Console      Empty row
            BREAK
        END

        ${row_num}=         Evaluate    ${row_num} + 1

        ${status}=         Run Keyword And Return Status
        ...                Should Contain Match    ${row}    ${row_identifier_arg}

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

    # Create a suite variable from each value in the row
    FOR    ${key_value_tuple}    IN    &{row_dict}
        IF    "${key_value_tuple}[1]" != ""
            Set Suite Variable          ${${key_value_tuple}[0]}    ${key_value_tuple}[1]
        END
    END

Read Excel Data
    [Documentation]     Reads all rows from an excel sheet to a list of dictionaries. Dictionary keys are values
    ...                 from the first row.
    [Arguments]         ${file_arg}    ${sheet_name_arg}

    # Open existing workbook
    ${document}=        Open Excel Document    ${file_arg}    excel

    ${header_row}=	    Read Excel Row	    row_num=1    sheet_name=${sheet_name_arg}


    @{list_of_rows}=    Create List
    ${row_num}=         Set Variable    ${2}

    # Set the limit to what you think is the max amount of rows
    WHILE    True    limit=200
        ${row}=	            Read Excel Row	    row_num=${row_num}    sheet_name=${sheet_name_arg}

        # Break from loop when row is empty, we have reached the end of the sheet
        IF    "${row}[0]" == "None"
            Log To Console      Empty row
            BREAK
        END

        ${row_num}=         Evaluate    ${row_num} + 1

        &{row_dict}=        Create Dictionary
        FOR    ${index}    ${header}    IN ENUMERATE    @{header_row}
            Set To Dictionary    ${row_dict}    ${header}    ${row}[${index}]
        END
        
        Append To List      ${list_of_rows}    ${row_dict}
    END

    Close All Excel Documents

    RETURN               ${list_of_rows}
