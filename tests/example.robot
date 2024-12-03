*** Settings ***
Library           PywinautoLibrary

*** Test Cases ***
Open Notepad, Type Text and Save
    Launch Application    notepad.exe
    Get Dialog From Regex   .*Untitled
    Type Text            Edit    Hello, World!
    Menu Select   File->Save As
    Get Dialog    Save As
    Type Text    edit1    test.txt
    Click    Save
    Close Application
