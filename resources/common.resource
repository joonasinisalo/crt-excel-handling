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
