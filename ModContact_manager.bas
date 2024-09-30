Attribute VB_Name = "ModContact_manager"
Option Explicit
Public ContactName, ContactAddress, City, State, PhoneNumber, Email As String


Sub AddnewContact()
    Range("E5,E7,E9,H5,H7,H9").ClearContents
    

End Sub

Sub SaveContact()

ContactName = Range("E5")
City = Range("E7")
PhoneNumber = Range("E9")
ContactAddress = Range("H5")
State = Range("H7")
Email = Range("H9")

Dim NextRecord As Range

Set NextRecord = ActiveSheet.Range("D1048576").End(xlUp).Offset(1)

Dim a, b As Range

Set a = ActiveSheet.Range("E5")
Set b = ActiveSheet.Range("D1048576").End(xlUp)

Dim icell As Range

For Each icell In Range("D13", Range("D13").End(xlDown)).Cells
    If (icell = a) Then MsgBox ("This contact name already exists. Click the button 'View Contact' to see the details.")
    If (icell = a) Then Exit Sub
Next icell
    
NextRecord = ContactName
NextRecord.Offset(, 1) = ContactAddress
NextRecord.Offset(, 2) = City
NextRecord.Offset(, 3) = State
NextRecord.Offset(, 4) = PhoneNumber
NextRecord.Offset(, 5) = Email

End Sub

Sub ViewContact()

Dim icell As Range

ContactName = Range("E5")
City = Range("E7")
PhoneNumber = Range("E9")
ContactAddress = Range("H5")
State = Range("H7")
Email = Range("H9")

    For Each icell In Range("D13", Range("I13").End(xlDown)).Cells
        If icell = ContactName Then
            Range("H5") = icell.Offset(, 1)
            Range("E7") = icell.Offset(, 2)
            Range("H7") = icell.Offset(, 3)
            Range("E9") = icell.Offset(, 4)
            Range("H9") = icell.Offset(, 5)
        ElseIf icell = PhoneNumber Then
            Range("E5") = icell.Offset(, -4)
            Range("E7") = icell.Offset(, -3)
            Range("H5") = icell.Offset(, -2)
            Range("H7") = icell.Offset(, -1)
            Range("H9") = icell.Offset(, 1)
        ElseIf icell = Email Then
            Range("E5") = icell.Offset(, -5)
            Range("H5") = icell.Offset(, -4)
            Range("E7") = icell.Offset(, -3)
            Range("H7") = icell.Offset(, -2)
            Range("E9") = icell.Offset(, -1)
        End If
    Next icell
    
End Sub

Sub DeleteContact()

If MsgBox("Are you sure, you want to delete this contact?", vbYesNo + vbQuestion, "Delete Contact?") = vbNo Then Exit Sub

Dim icell As Range

ContactName = Range("E5")

For Each icell In Range("D13", Range("D13").End(xlDown)).Cells
    If icell = ContactName Then
        icell.ClearContents
        icell.Offset(, 1).ClearContents
        icell.Offset(, 2).ClearContents
        icell.Offset(, 3).ClearContents
        icell.Offset(, 4).ClearContents
        icell.Offset(, 5).ClearContents
    End If
Next icell

Range(Range(("D13")).End(xlDown).Offset(1), Range("D13").End(xlDown).Offset(1, 5)).Delete

End Sub

Sub EditContact()

Dim icell As Range

ContactName = Range("E5")
City = Range("E7")
PhoneNumber = Range("E9")
ContactAddress = Range("H5")
State = Range("H7")
Email = Range("H9")

If MsgBox("Are you sure, you want to make changes?", vbYesNo + vbQuestion, "Make Changes?") = vbNo Then Exit Sub

For Each icell In Range("I13", Range("I13").End(xlDown)).Cells  'changes based on email'
    If icell = Email Then
        icell.Offset(, -1) = PhoneNumber
        icell.Offset(, -2) = State
        icell.Offset(, -3) = City
        icell.Offset(, -4) = ContactAddress
        icell.Offset(, -5) = ContactName
    End If
Next icell

For Each icell In Range("D13", Range("D13").End(xlDown)).Cells  'changes based on email'
    If icell = ContactName Then
        icell.Offset(, 1) = ContactAddress
        icell.Offset(, 2) = City
        icell.Offset(, 3) = State
        icell.Offset(, 4) = PhoneNumber
        icell.Offset(, 5) = Email
    End If
Next icell

End Sub

Sub PreviousContact()

Dim icell As Range

ContactName = Range("E5")
City = Range("E7")
PhoneNumber = Range("E9")
ContactAddress = Range("H5")
State = Range("H7")
Email = Range("H9")

    For Each icell In Range("D14", Range("D14").End(xlDown)).Cells
        If icell = ContactName Then
            Range("E5") = icell.Offset(-1)
            Range("H5") = icell.Offset(-1, 1)
            Range("E7") = icell.Offset(-1, 2)
            Range("H7") = icell.Offset(-1, 3)
            Range("E9") = icell.Offset(-1, 4)
            Range("H9") = icell.Offset(-1, 5)
        End If
    Next icell
            
End Sub

Sub NextContact()

Dim icell As Range

ContactName = Range("E5")
City = Range("E7")
PhoneNumber = Range("E9")
ContactAddress = Range("H5")
State = Range("H7")
Email = Range("H9")

    For Each icell In Range("D13", Range("D13").End(xlDown)).Cells
        If icell = ContactName Then
            Range("E5") = icell.Offset(1)
            Range("H5") = icell.Offset(1, 1)
            Range("E7") = icell.Offset(1, 2)
            Range("H7") = icell.Offset(1, 3)
            Range("E9") = icell.Offset(1, 4)
            Range("H9") = icell.Offset(1, 5)
        End If
    Next icell

End Sub
