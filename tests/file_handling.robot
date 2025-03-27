*** Settings ***
Library                 QWeb
Library                 OperatingSystem
Library                 Process
Library                 String
Library                 DateTime
Library                 ${CURDIR}/../libraries/GitOperations.py
Suite Setup             Open Browser    about:blank    chrome
Suite Teardown          Close All Browsers


*** Test Cases ***
Create And Save File
    ${today}=           Get Current Date    result_format=%Y-%m-%d

    ${file}=            Set Variable    ${CURDIR}/../files/date_file.txt
    Log                 ${file}    console=True

    Append To File      ${file}    ${today}\n

    ${file_content}=    Get File    ${file}
    Log                 ${file_content}    console=True

    Commit And Push     ${file}     main
