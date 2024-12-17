*** Settings ***
Library    OperatingSystem
Library    String
Library    RPA.FileSystem
Library    RPA.Windows
Library    RPA.Desktop
Library    Dialogs
Library    DateTime
Library    Collections
Library    tasks.py

*** Variables ***
${company}    IL
${branch}    07
${invalid_login}    False
# ${invoice_folder}    ${CURDIR}/Westport PDF Invoice

*** Test Cases ***
Westport PDF Invoice To Sovy
    # Prompt user password for SOVY
    ${password}=    Set Variable    ${EMPTY}
    WHILE    '${password}' == ''
        ${password}=    Get Value From User    Input your SOVY password:    hidden=yes
    END
    ${config}    Read File    config.txt
    ${config}    Split String    ${config}    \n
    Start Sovy    ${config}[0]    ${password}

    # Close Sovy
    Run Keyword If    ${invalid_login} == True    RPA.Windows.Click    name:Cancel
    Run Keyword If    ${invalid_login} == False    Stop Sovy
    
*** Keywords ***
Read All Pdf Files In Directory
    [Arguments]    ${path}
    ${files}    OperatingSystem.List Files In Directory    ${path}
    ${pdf_files}    Create List
    FOR    ${file}    IN    @{files}
        Run Keyword If    "${file.lower().endswith('.pdf')}"    Append To List    ${pdf_files}    ${file}
    END
    RETURN    ${pdf_files}

Start Sovy
    [Arguments]    ${invoice_folder}    ${password}
    Windows Search    Sovy Logistics Solution (UAT)
    Sleep    20
    RPA.Windows.Click    name:PasswordTextEdit
    RPA.Desktop.Type Text    ${password}
    RPA.Windows.Click    automationid:btnOk
    Sleep    1
    # "Information" window pops up means invalid login (password incorrect).
    ${fail_login}    Run Keyword And Ignore Error    RPA.Windows.Click    name:Information
    # Check if the click on "Information" was successful (element found)
    Run Keyword If    '${fail_login[0]}' == 'PASS'    Click OK and Proceed    Invalid Login!
    Run Keyword If    '${fail_login[0]}' == 'PASS'    Set Global Variable   ${invalid_login}    True
    Run Keyword If    '${fail_login[0]}' == 'FAIL'    Continue Sovy Operation    ${invoice_folder}

Continue Sovy Operation
    [Arguments]    ${invoice_folder}
    Handle Dropdown    ${company}
    RPA.Desktop.Press Keys         TAB
    Sleep    1
    Handle Dropdown    ${branch}
    RPA.Windows.Click    automationid:btnOk
    Sleep    2
    Control Window    [IL/07] Sovy Logistics Solutions Version 5.0.24.41013 (UAT)
    @{files}    Read All Pdf Files In Directory    ${invoice_folder}
    FOR    ${filename}    IN    @{files}
        ${invoice_no}    ${date}    ${quantity}    ${rate}    ${container_no}    Extract Pdf Invoice Info    ${invoice_folder}    ${filename}
        RPA.Windows.Click    name:Operation
        Sleep    2 
        # Navigate to select "Job Entry Listing"
        # TAB 2 -> 4 (user)
        RPA.Desktop.Press Keys         TAB
        RPA.Desktop.Press Keys         TAB
        # RPA.Desktop.Press Keys         TAB
        # RPA.Desktop.Press Keys         TAB
        RPA.Desktop.Press Keys         ENTER
        Sleep    10
        # Navigate to the "From" field 
        RPA.Desktop.Press Keys         TAB
        RPA.Desktop.Press Keys         TAB
        RPA.Desktop.Press Keys         TAB
        RPA.Desktop.Type Text    ${date}
        Sleep    2
        # Navigate to the "Container No" field 
        RPA.Desktop.Press Keys         TAB
        RPA.Desktop.Press Keys         TAB
        RPA.Desktop.Press Keys         TAB
        RPA.Desktop.Press Keys         TAB
        RPA.Desktop.Type Text    ${container_no}
        Sleep    2
        # Click the "Retrieve" button
        RPA.Windows.Click    name:Retrieve
        Sleep    5
        
        # "Information" window pops up means no record was found.
        ${result}    Run Keyword And Ignore Error    RPA.Windows.Click    name:Information
    
        # Check if the click on "Information" was successful (element found)
        Run Keyword If    '${result[0]}' == 'PASS'    Click OK and Proceed    No Record Found!
        Run Keyword If    '${result[0]}' == 'FAIL'    Get Job No and Fill Vendor Invoice    ${invoice_no}    ${quantity}    ${rate}
    END    

Handle Dropdown
    [Arguments]    ${attribute_value}
    RPA.Desktop.Press Keys         alt    down
    FOR    ${i}    IN RANGE    0    20
        ${row_name}=    Set Variable    name:"Row ${i}"
        ${row_value}=    RPA.Windows.Get Value    ${row_name}
        ${code_in_row}=    Evaluate    str("${row_value}").split(";")[0]
        Run Keyword If    '${code_in_row}' == '${attribute_value}'    RPA.Windows.Click    ${row_name}
        Run Keyword If    '${code_in_row}' == '${attribute_value}'    Exit For Loop
    END

Click OK and Proceed
    [Arguments]    ${message}
    RPA.Windows.Click    name:OK
    Log To Console    ALERT: ${message}
    Sleep    2

Get Job No and Fill Vendor Invoice
    [Arguments]    ${invoice_no}    ${quantity}    ${rate}
    # Select the first record, which is also the latest one
    # RPA.Windows.Click    name:"Job No. row 1"
    RPA.Desktop.Click    coordinates:65,420    #65,420-> 65,280(for user), depends on display resolution
    Sleep    20
    ${job_no} =  Get Value   automationid:edtJobNo
    Sleep    2

    RPA.Windows.Click    name:"Billing & Costing"
    Sleep    2 
    # Navigate to select "Vendor Invoice"
    # TAB 6 -> 10 (user)
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    # RPA.Desktop.Press Keys         TAB
    # RPA.Desktop.Press Keys         TAB
    # RPA.Desktop.Press Keys         TAB
    # RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         ENTER
    Sleep    5
    
    # Fill in a new vendor invoice
    RPA.Windows.Click    name:New
    Sleep    1
    # Vendor
    RPA.Desktop.Click    coordinates:200,170    #200,170-> 200,113(for user), depends on display resolution
    Sleep    1
    RPA.Desktop.Type Text    W0019
    Sleep    1
    # Vendor Ref
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Type Text    ${invoice_no}
    RPA.Desktop.Press Keys         TAB
    Sleep    1
    # "Error" window pops up means vendor ref already exist.
    ${result}    Run Keyword And Ignore Error    RPA.Windows.Click    name:Error

    # Check if the click on "Information" was successful (element found)
    Run Keyword If    '${result[0]}' == 'PASS'    Click OK and Proceed    Vendor Ref already exist!
    Run Keyword If    '${result[0]}' == 'FAIL'    Continue Filling Invoice    ${job_no}    ${quantity}    ${rate}

Continue Filling Invoice
    [Arguments]    ${job_no}    ${quantity}    ${rate}
    # Invoice Date – Today’s date
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    ${today_date}=    Get Current Date    result_format=%d/%m/%Y
    RPA.Desktop.Type Text    ${today_date}
    Sleep    1
    # Job No.
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Type Text    ${job_no}
    Sleep    1
    # Charge Code
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Type Text    FV5135
    Sleep    1
    # Quantity
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Type Text    ${quantity}
    Sleep    1
    # Price
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Press Keys         TAB
    RPA.Desktop.Type Text    ${rate}
    Sleep    1
    #Save the record
    RPA.Windows.Click    name:Save
    Sleep    2
    RPA.Windows.Click    name:OK
    Sleep    2

Stop Sovy
    RPA.Windows.Click    name:Item
    Sleep    1
    RPA.Windows.Click    name:Yes
    Sleep    1
    # If the save changes window pop up, close the window
    Run Keyword And Ignore Error    RPA.Windows.Click    name:No
    
