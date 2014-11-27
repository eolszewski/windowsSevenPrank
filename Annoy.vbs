'This is a script which will wait between one and three hours to prompt the user and then eject their cd drives 
'If you would like this script to run on startup of a computer, 
'it must be placed in the following directory for Windows 7: "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
'I have included a script called moveFile.vbs which will move the file to that directory, but if you are not an administrator,
'you will have to do this manually. You then need to click on this file to execute it.

'Getting access to the CD drives
Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROMs = oWMP.cdromCollection
do
    'Random function between one and three hours
    Dim max,min
    max=10800000
    min=3600000

    'Waiting that amount of time before executing code
    wscript.sleep Int((max-min+1)*Rnd+min)
    
    'Prompting user for a cookie
    x=msgbox("Please insert cookie." ,16, "Cookie Needed.") 
    'Ejecting or closing all cd drives
    For i = 0 to colCDROMs.Count - 1
        colCDROMs.Item(i).Eject
    Next
loop