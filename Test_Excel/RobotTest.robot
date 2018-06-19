*** Settings ***
Suite Setup       Open Browser    http://newtours.demoaut.com/    chrome
Suite Teardown    Close Browser
Test Setup        Set Global Variable    ${Path}    D:/Robot_Scipts/Test_Excel/
Library           Selenium2Library
Library           ExcelLibrary
Library           String
Library           DateTime
Library           OperatingSystem
Library           Collections

*** Variables ***

*** Test Cases ***
Main
    Set Screenshot Directory    ${Path}
    Count All Data From Excel
    : FOR    ${iterarion}    IN RANGE    1    ${AllRow}
    \    Get Data From Excel    ${iterarion}
    \    Login
    \    Comment    Flight Finder
    \    Comment    Select Flight
    \    Comment    Book a Flight
    \    Comment    Flight Confirmation
    \    Logout    ${iterarion}

*** Keyword ***
Count All Data From Excel
    Open Excel    D:\\Robot_Scipts\\Test_Excel\\DataTest.xls
    @{SheetName}=    Get Sheet Names
    ${AllRow}=    Get Row Count    @{SheetName}[0]
    Set Suite Variable    @{SheetName}
    Set Suite Variable    ${AllRow}

Get Data From Excel
    [Arguments]    ${iterarion}
    @{CusFirstName}=    Create List
    @{CusLastName}=    Create List
    ${User}=    Read Cell Data By Coordinates    Sheet1    0    ${iterarion}
    ${Password}=    Read Cell Data By Coordinates    Sheet1    1    ${iterarion}
    ${FromPort}=    Read Cell Data By Coordinates    Sheet1    2    ${iterarion}
    ${ToPort}=    Read Cell Data By Coordinates    @{SheetName}[0]    3    ${iterarion}
    ${Service}=    Read Cell Data By Coordinates    @{SheetName}[0]    4    ${iterarion}
    ${Depart}=    Read Cell Data By Coordinates    @{SheetName}[0]    5    ${iterarion}
    ${Returen}=    Read Cell Data By Coordinates    @{SheetName}[0]    6    ${iterarion}
    ${Name1}=    Read Cell Data By Coordinates    @{SheetName}[0]    7    ${iterarion}
    ${LastName1}=    Read Cell Data By Coordinates    @{SheetName}[0]    8    ${iterarion}
    ${Name2}=    Read Cell Data By Coordinates    @{SheetName}[0]    9    ${iterarion}
    ${LastName2}=    Read Cell Data By Coordinates    @{SheetName}[0]    10    ${iterarion}
    Append To List    ${CusFirstName}    ${Name1}
    Append To List    ${CusFirstName}    ${Name2}
    Append To List    ${CusLastName}    ${LastName1}
    Append To List    ${CusLastName}    ${LastName2}
    Set Suite Variable    ${User}
    Set Suite Variable    ${Password}
    Set Suite Variable    ${FromPort}
    Set Suite Variable    ${ToPort}
    Set Suite Variable    ${Service}
    Set Suite Variable    ${Depart}
    Set Suite Variable    ${Returen}
    Set Suite Variable    ${CusFirstName}
    Set Suite Variable    ${CusLastName}
    Comment    Set Suite Variable    ${Name1}
    Comment    Set Suite Variable    ${LastName1}
    Comment    Set Suite Variable    ${Name2}
    Comment    Set Suite Variable    ${LastName2}
    Put String To Cell    @{SheetName}[0]    11    ${iterarion}    Passed
    Save Excel    D:\\Robot_Scipts\\Test_Excel\\TestResult.xls

Login
    Wait Until Element Is Visible    name=userName    timeout=30
    Input Text    name=userName    ${User}
    Input Text    name=password    ${Password}
    Capture Page Screenshot
    Click Element    name=login

Flight Finder
    Wait Until Page Contains    Flight Finder    timeout=30
    Select Radio Button    tripType    oneway
    Select From List By Value    name=passCount    2
    Select From List By Value    name=fromPort    ${FromPort}
    Select From List By Value    name=toPort    ${ToPort}
    Select Radio Button    servClass    ${Service}
    Capture Page Screenshot
    Click Element    name=findFlights

Select Flight
    Wait Until Element Is Visible    xpath=//img[@src="/images/masts/mast_selectflight.gif"]    timeout=30
    Select Radio Button    outFlight    ${Depart}
    Select Radio Button    inFlight    ${Returen}
    Capture Page Screenshot
    Click Element    name=reserveFlights

Book a Flight
    Wait Until Element Is Visible    xpath=//img[@src="/images/masts/mast_book.gif"]    timeout=30
    : FOR    ${i}    IN RANGE    0    2
    \    Comment    Run Keyword If    ${i}==0    Input Text    name=passFirst${i}    ${Name1}
    \    ...    ELSE    Input Text    name=passFirst${i}    ${Name2}
    \    Comment    Run Keyword If    ${i}==0    Input Text    name=passLast${i}    ${LastName1}
    \    ...    ELSE    Input Text    name=passLast${i}    ${LastName2}
    \    Input Text    name=passFirst${i}    @{CusFirstName}[${i}]
    \    Input Text    name=passLast${i}    @{CusLastName}[${i}]
    Input Text    name=creditnumber    1234567890123456
    Select Checkbox    name=ticketLess
    Capture Page Screenshot
    Click Element    name=buyFlights

Flight Confirmation
    Wait Until Page Contains    Flight Confirmation    timeout=30
    ${StrText}=    Get Text    xpath=/html[1]/body[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/table[1]/tbody[1]/tr[4]/td[1]/table[1]/tbody[1]/tr[1]/td[2]/table[1]/tbody[1]/tr[5]/td[1]/table[1]/tbody[1]/tr[3]/td[1]/font[1]
    ${StrText2}=    Get Text    xpath=/html[1]/body[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/table[1]/tbody[1]/tr[4]/td[1]/table[1]/tbody[1]/tr[1]/td[2]/table[1]/tbody[1]/tr[5]/td[1]/table[1]/tbody[1]/tr[12]/td[1]/table[1]/tbody[1]
    Click Element    xpath=/html[1]/body[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/table[1]/tbody[1]/tr[4]/td[1]/table[1]/tbody[1]/tr[1]/td[2]/table[1]/tbody[1]/tr[5]/td[1]/table[1]/tbody[1]/tr[12]/td[1]/table[1]/tbody[1]
    Capture Page Screenshot

Logout
    [Arguments]    ${iterarion}
    Click Link    mercurysignoff.php
    Wait Until Element Is Visible    xpath=//img[@src="/images/masts/mast_signon.gif"]    timeout=30
