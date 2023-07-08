Option Explicit
do
Call Open_Close_CD()

Call Pause(10)

Call Open_Close_CD()
loop
'****************************************
Sub Open_Close_CD()
Dim oWMP,colCDROMs,i
Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROMs = oWMP.cdromCollection
if colCDROMs.Count >= 1 then
    colCDROMs.Item(i).Eject
End If
End Sub
'****************************************
Sub Pause(NSeconds)
    Wscript.Sleep(NSeconds*1000)
End Sub
'****************************************