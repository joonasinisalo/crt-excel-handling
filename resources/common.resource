*** Settings ***
Library                 QForce
Library                 String


*** Variables ***
${BROWSER}              chrome


*** Keywords ***
Setup Browser
    Open Browser        about:blank     ${BROWSER}
    Evaluate            random.seed()    random    # initialize random generator

End Suite
    Close All Browsers

Login
    [Arguments]           ${sf_instance_url}=${login_url}    ${sf_username}=${username}   ${sf_password}=${password}  
    GoTo                  ${sf_instance_url}
    TypeText              Username    ${sf_username}    delay=1
    TypeSecret            Password    ${sf_password}
    ClickText             Log In
