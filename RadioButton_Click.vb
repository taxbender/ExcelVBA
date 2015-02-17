Sub RadioButton_Click()

Dim ctrl As MSForms.Control
Dim dict As Object

Set dict = CreateObject("Scripting.Dictionary")


'## Iterate the controls, and add the GroupName and Button.Name
'  to a Dictionary object if the button is True.
'  use the GroupName as the unique identifier/key, and control name as the value

For Each ctrl In Me.Controls
    If TypeName(ctrl) = "OptionButton" And ctrl.Value = True Then
        dict(ctrl.GroupName) = ctrl.Name
    End If
Next

'## Now, to call on the values you simply refer to the dictionary by the GroupName, so:

Debug.Print dict("Family A")
Debug.Print dict("Family B")

Set dict = Nothing

End Sub
