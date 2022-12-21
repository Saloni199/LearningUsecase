*** Settings ***
Documentation       Extraction of Learning and send email with excel file.

Library             RPA.Browser.Selenium    auto_close=${True}
Library             RPA.Excel.Files
Library             RPA.Robocorp.Vault
Library             RPA.Outlook.Application



*** Variables ***
${Url}=
...                             https://performancemanager10.successfactors.com/sf/home?bplte_company=yashtechnoP&_s.crb=LEYx2FnSst29S25zpSb6k5f8rfst45lJSDPxYg6cOW8%253d=${TRUE}
${Workbook}=                    Infogram.xlsx
${TotalLearning} =              3
${TotalLearningHistory} =       3
${recipients}=                 abc@yash.com



*** Tasks ***
Extraction of Learning and send email with excel file.
     TRY
        Login To Infogram
        Navigate to Learning Course
        Extraction of Learning Course and Write into excel(Infogram.xlsx)
        Adding New Worksheet To Infogram Excel
        Navigate to Learning History on Infogram
        Extraction Of Learning History and Write into excel(Infogram.xlsx)
  
        Send Email with attachments
     EXCEPT
       exception
    END


*** Keywords ***
Login To Infogram
    Open Available Browser
    ...    ${Url}
    ...    browser_selection=chrome
    ...    maximized=TRUE
    Open Workbook    ${Workbook}
    ${secret}=    Get Secret    credentials
    Input Text    UserName    ${secret}[username]
    Sleep    1s
    Input Password    Password    ${secret}[password]    ${TRUE}
    Sleep    2s
    Click Element    id:submitButton

Navigate to Learning Course
    Sleep    2s
    Click Link    xpath=//*[@id="content"]/div/div[2]/div/section/ul/li[4]/ui5-busy-indicator/a
    Sleep    20s
    Click Element When Visible    class:learnerToDoGroupItemOverDueStyle
  
    Click Element When Visible    xpath=//tbody[3]/tr/th/a/span[2]
    Sleep    5s

Extraction of Learning Course and Write into excel(Infogram.xlsx)
    FOR    ${counter}    IN RANGE    1    ${TotalLearning}
        ${row}=    Find Empty Row
        ${Course}=    Get Text
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div[2]/div/div/div[2]/div/div/table/tbody[6]/tr[${counter}]/td/div/div[3]/div[2]/div/div[1]/span[1]/a
        Sleep    2s
        Mouse Over
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div[2]/div/div/div[2]/div/div/table/tbody[6]/tr[${counter}]/td/div/div[3]/div[2]/div/div[1]/span[1]/a
        Sleep    5s
        Click Element When Visible
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div[2]/div/div/div[2]/div/div/table/tbody[6]/tr[${counter}]/td/div/div[3]/div[2]/div/div[1]/span[2]/a
        Is Element Visible
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div[2]/div/div/div[2]/div/div/table/tbody[6]/tr[${counter}]/td/div/div[4]/div/dl/dd[1]
        Sleep    5s
        ${Desc}=    Get Text
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div[2]/div/div/div[2]/div/div/table/tbody[6]/tr[${counter}]/td/div/div[4]/div/dl/dd[1]
        Sleep    5s
        ${Type}=    Get Text
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div[2]/div/div/div[2]/div/div/table/tbody[6]/tr[${counter}]/td/div/div[4]/div/dl/dd[2]
        Sleep    2s
        ${Credit}=    Get Text
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div[2]/div/div/div[2]/div/div/table/tbody[6]/tr[${counter}]/td/div/div[4]/div/div/ul/li/div/span[1]/span[1]
        Set Active Worksheet    Learnings
        Set Cell Value    ${row}    A    ${Course}
        Set Cell Value    ${row}    B    ${Desc}
        Set Cell Value    ${row}    C    ${Type}
        Set Cell Value    ${row}    D    ${Credit}
        

        ${DescBool}=    Is Element Visible
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div[2]/div/div/div[2]/div/div/table/tbody[6]/tr[${counter}]/td/div/div[4]/div/dl/dd[1]
        IF    ${DescBool} == False    Set Cell Value    ${row}    B    NA
        ${CourseBool}=    Is Element Visible
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div[2]/div/div/div[2]/div/div/table/tbody[6]/tr[${counter}]/td/div/div[3]/div[2]/div/div[1]/span[1]/a
        IF    ${CourseBool} == False    Set Cell Value    ${row}    A    NA
        ${TypeBool}=    Is Element Visible
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div[2]/div/div/div[2]/div/div/table/tbody[6]/tr[${counter}]/td/div/div[4]/div/dl/dd[2]
        IF    ${TypeBool} == False    Set Cell Value    ${row}    C    NA
        ${TCreditBool}=    Is Element Visible
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div[2]/div/div/div[2]/div/div/table/tbody[6]/tr[${counter}]/td/div/div[4]/div/dl/dd[2]
        IF    ${TypeBool} == False    Set Cell Value    ${row}    D    NA
        Save Workbook    Infogram.xlsx
    END

Send Email with attachments
    RPA.Outlook.Application.Open Application
    Send Message    recipients=${recipients}
    ...    subject=Hello
    ...    body=Please see the attached file for Infogram.
    ...    attachments=${Workbook}
    RPA.Outlook.Application.Quit Application


Navigate to Learning History on Infogram
    Click Element When Visible
    ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[2]/div/div/div[3]/div/div/div/div[1]/h2/a/span
    Sleep    5s

Extraction Of Learning History and Write into excel(Infogram.xlsx)
    FOR    ${counter}    IN RANGE    1    ${TotalLearningHistory}
        ${CompletionDate}=    Get Text
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[3]/div/div[2]/div/div/div/div/table/tbody/tr[${counter}]/td[1]/span
        Sleep    2s
        ${CourseName}=    Get Text
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[3]/div/div[2]/div/div/div/div/table/tbody/tr[${counter}]/td[2]/div/div[2]/span/a
        Sleep    2s
        ${Status}=    Get Text
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[3]/div/div[2]/div/div/div/div/table/tbody/tr[${counter}]/td[3]/span

        Set Active Worksheet    Learning History
        ${row1}=    Find Empty Row    Learning History
        Set Cell Value    ${row1}    A    ${CompletionDate}
        Set Cell Value    ${row1}    B    ${CourseName}
        Set Cell Value    ${row1}    C    ${Status}

        ${CompletionDateBool}=    Is Element Visible
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[3]/div/div[2]/div/div/div/div/table/tbody/tr[${counter}]/td[1]/span
        IF    ${CompletionDateBool} == False
            Set Cell Value    ${row1}    A    NA
        END
        ${CourseNameBool}=    Is Element Visible
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[3]/div/div[2]/div/div/div/div/table/tbody/tr[${counter}]/td[2]/div/div[2]/span/a
        IF    ${CourseNameBool} == False    Set Cell Value    ${row1}    B    NA
        ${StatusBool}=    Is Element Visible
        ...    xpath=/html/body/div[2]/div[2]/div/div[4]/div[3]/div/div[2]/div/div/div/div/table/tbody/tr[${counter}]/td[3]/span
        IF    ${StatusBool} == False    Set Cell Value    ${row1}    C    NA

        Save Workbook    Infogram.xlsx
    END

Adding New Worksheet To Infogram Excel
    ${CheckSheet}=    Worksheet Exists    Learning History
   
    IF    ${CheckSheet} == True
        Set Active Worksheet    Learning History
        Set Worksheet Value    1    A    CompletionDates
        Set Worksheet Value    1    B    CourseName
        Set Worksheet Value    1    C    Status
        Save Workbook    Infogram.xlsx
    ELSE
        Create Worksheet    Learning History
        Set Worksheet Value    1    A    CompletionDates
        Set Worksheet Value    1    B    CourseName
        Set Worksheet Value    1    C    Status
        Save Workbook    Infogram.xlsx
    END

exception
    Log    Exception occ check logs.
