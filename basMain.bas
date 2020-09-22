Attribute VB_Name = "basMain"
Option Explicit
'------------------------------------------------
' This example (new to vbXML 0.02) shows off the
' power of the three new functions.  If you still
' need or want the example for 0.01 then just
' email me and I will send it asap.
'             akartanis@akguild.com
'
' I think I will make one final example with
' the 0.03 release the shows off all of the power
' of vbXML all in one package.  This will allow
' for easy learning of the vbXML way of life ;)
' The example will probably be for skinning
' something, which is why I plan to add some
' functions in 0.03 that will make skinning
' easier (if a specific format is followed in the
' file to be skinned)
'
' See also: Readme.txt
'------------------------------------------------
Public XML As New vbXML

Public cCount As String

Public Name As String
Public Phone As String
Public Email As String
Public Street As String
Public City As String
Public State As String
Public Zip As String
Public Notes As String

Public Sub LoadContactList()
Dim i As Long

With frmMain
For i = 0 To XML.NodeCount("/contacts") - 1
    .lvNames.ListItems.Add , , XML.GetChildName("/contacts", i)
    .lvNames.ListItems(i + 1).SubItems(1) = XML.ReadNode("/contacts/" & XML.GetChildName("/contacts", i) & "/phone")
Next i
End With
End Sub

Public Sub LoadContactInfo(strName As String)
Dim i As Long

With frmMain
For i = 0 To XML.NodeCount("/contacts") - 1
    If XML.GetChildName("/contacts", i) <> strName Then GoTo Nexti:
    .txtAddress.Text = XML.ReadNode("/contacts/" & XML.GetChildName("/contacts", i) & "/street")
    .txtCity = XML.ReadNode("/contacts/" & XML.GetChildName("/contacts", i) & "/city")
    .txtEmail = XML.ReadNode("/contacts/" & XML.GetChildName("/contacts", i) & "/email")
    .txtInfo = XML.ReadNode("/contacts/" & XML.GetChildName("/contacts", i) & "/notes")
    .txtName = XML.GetChildName("/contacts", i) & " " & XML.ReadNode("/contacts/" & XML.GetChildName("/contacts", i) & "/lname")
    .txtPhone = XML.ReadNode("/contacts/" & XML.GetChildName("/contacts", i) & "/phone")
    .txtState = XML.ReadNode("/contacts/" & XML.GetChildName("/contacts", i) & "/state")
    .txtZip = XML.ReadNode("/contacts/" & XML.GetChildName("/contacts", i) & "/zip")
Nexti:
Next i
End With
End Sub

Public Sub SaveContactInfo(strName As String)
Dim i As Long
Dim cName As String

With frmMain
If XML.NodeCount("/contacts") = 0 Then GoTo NewNode:
For i = 0 To XML.NodeCount("/contacts") - 1
    cName = XML.GetChildName("/contacts", i)
    If cName = strName Then
        XML.WriteNode "/contacts/" & cName & "/street", .txtAddress
        XML.WriteNode "/contacts/" & cName & "/city", .txtCity
        XML.WriteNode "/contacts/" & cName & "/email", .txtEmail
        XML.WriteNode "/contacts/" & cName & "/notes", .txtInfo
        XML.WriteNode "/contacts/" & cName & "/lname", Right(.txtName, Len(.txtName) - (Len(cName) + 1))
        XML.WriteNode "/contacts/" & cName & "/phone", .txtPhone
        XML.WriteNode "/contacts/" & cName & "/state", .txtState
        XML.WriteNode "/contacts/" & cName & "/zip", .txtZip
        XML.Save App.Path & "\contact.xml"
        XML.OpenXML App.Path & "\contact.xml"
    Else
NewNode:
        XML.MakeNode "/contacts", strName
        XML.MakeNode "/contacts/" & strName, "street", .txtAddress
        XML.MakeNode "/contacts/" & strName, "city", .txtCity
        XML.MakeNode "/contacts/" & strName, "email", .txtEmail
        XML.MakeNode "/contacts/" & strName, "notes", .txtInfo
        XML.MakeNode "/contacts/" & strName, "lname", Right(.txtName, Len(.txtName) - (Len(cName) + 1))
        XML.MakeNode "/contacts/" & strName, "phone", .txtPhone
        XML.MakeNode "/contacts/" & strName, "state", .txtState
        XML.MakeNode "/contacts/" & strName, "zip", .txtZip
        XML.Save App.Path & "\contact.xml"
        XML.OpenXML App.Path & "\contact.xml"
        .lvNames.ListItems.Add , , strName
        .lvNames.ListItems(.lvNames.ListItems.Count).SubItems(1) = .txtPhone
    End If
Next i
End With
End Sub
