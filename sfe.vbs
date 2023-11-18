Option Explicit

Dim objWshShell

Dim intN                                                             
Dim intO                                                            
Dim intO1                                                      
Dim intN1                                                             


Set objWshShell = WScript.CreateObject("WScript.Shell")

intN    = 10000                                                  
intO  = 100                                                         

intO1 = Timer                                                   

Do
    
    intN1 = objWshShell.Popup( _
        "Made" & vbCrLf & "by" & vbCrLf & "Noon4ik", _
        intN - (Timer - intO1), _
        "Noon_ms", _
        vbOKOnly + vbInformation + &H40000)
    
  
  
Loop Until intN1 = -1 Or (Timer - intO1) > intO

Set objWshShell = Nothing

WScript.Quit 0