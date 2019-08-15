VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateSelect 
   Caption         =   "Template Selection"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7725
   OleObjectBlob   =   "TemplateSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TemplateSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim basePath As String
Dim MailTemplate As Outlook.MailItem

''' Action executed when the form becomes active.
''' Load the template files name in the drop down.
Private Sub UserForm_Activate()
    Dim path, home, file, e As String
    
    home = Environ("AppData")
    basePath = home + "\Microsoft\Templates\"
    file = basePath + "*.oft"
    
    e = Dir(file)
    While e <> ""
        Me.Template.AddItem (e)
        e = Dir()
    Wend
End Sub

''' Action executed when a template is selected.
''' Load the template and fill the parameters inputs.
Private Sub Template_Change()
    Dim selected, path As String
    Dim vSubject, vContent, params As Variant
    
    selected = Me.Template.Value
    path = basePath + selected
    
    Set MailTemplate = Application.CreateItemFromTemplate(path)
    
    vSubject = FindParams(MailTemplate.subject)
    vContent = FindParams(MailTemplate.HTMLBody)
    MergeArray vSubject, vContent
    params = Uniq(vSubject)
    SetIntput params
    
    Me.OpenBtn.Enabled = True
End Sub

''' Action executed when the open button is clicked.
''' Replace in the template the parameters with the input value.
Private Sub OpenBtn_Click()
    Dim mi As Outlook.MailItem: Set mi = MailTemplate
    Dim values As Object: Set values = grabParamsValue
    Dim keys As Variant: keys = values.keys
    Dim i As Integer
    For i = 0 To (values.Count - 1)
        Dim k As String: k = keys(i)
        Dim v As String: v = values(k)
        Dim exp As Object: Set exp = CreateObject("VBScript.RegExp")
        With exp
            .Pattern = "\{:" & k & "\}"
            .Global = True
        End With
        
        If v = "" Then
            v = "{" & k & "}"
        End If
        
        mi.subject = exp.Replace(mi.subject, v)
        mi.HTMLBody = exp.Replace(mi.HTMLBody, v)
    Next i
    Unload Me
    
    mi.Display True
End Sub

''' Find the parameters in the content.
''' content : The content in which the parameters needs to be found.
''' Returns a list of all the parameters in the content
Private Function FindParams(ByRef content As String) As String()
    Dim rst As Object
    Dim rtn() As String
    Dim regex As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "\{(?:&nbsp;)?:([^{}]+)\}"
        .Global = True
    End With
    Set rst = regex.Execute(content)
    
    ReDim rtn(0)
    For Each m In rst
        rtn(UBound(rtn)) = m.SubMatches(0)
        ReDim Preserve rtn(0 To (UBound(rtn) + 1))
    Next m
    
    If (UBound(rtn) > 0) Then
        ReDim Preserve rtn(0 To (UBound(rtn) - 1))
    End If
    
    FindParams = rtn
End Function

''' Merge two arrays
''' dst : The destination array.
''' src : The source array.
''' Returns the merged array.
Private Sub MergeArray(ByRef dst As Variant, ByRef src As Variant)
    For Each p In src
        ReDim Preserve dst(0 To (UBound(dst) + 1))
        dst(UBound(dst)) = p
    Next p
End Sub

''' Filter out duplicate keys from an array.
''' arr : The array to filter.
''' Returns the filtered array.
Private Function Uniq(ByRef arr) As Variant
    Dim rtn As Variant: ReDim rtn(0)
    
    For Each e In arr
        Dim add As Boolean: add = True
        
        For Each re In rtn
            If re = e Then
                add = False
                Exit For
            End If
        Next re
        
        If add Then
            rtn(UBound(rtn)) = e
            ReDim Preserve rtn(0 To (UBound(rtn) + 1))
        End If
    Next e
    
    ReDim Preserve rtn(0 To (UBound(rtn) - 1))
    Uniq = rtn
End Function

''' Create the inputs in the variable frame based on the detected parametes
''' params : The list of parameters in the mail.
Private Sub SetIntput(ByRef params As Variant)
    Dim padTB, padL, spacing, scroll As Integer
    padTB = 7
    padLR = 5
    spacing = 20
    
    Me.VariablesFrame.Controls.Clear

    Dim cnt As Integer: cnt = 0
    For Each p In params
        ' Make the label
        Dim lbl As MSForms.label
        Set lbl = Me.VariablesFrame.Controls.add("Forms.Label.1", p & "Lbl")
        lbl.Caption = p
        lbl.Top = (padTB + 4) + (cnt * spacing)
        lbl.Left = padLR
        
        ' Make the Textbox
        Dim tb As MSForms.TextBox
        Set tb = Me.VariablesFrame.Controls.add("Forms.TextBox.1", p)
        tb.Top = padTB + (cnt * spacing)
        tb.Left = 100
        tb.Width = Me.VariablesFrame.Width - (padLR + tb.Left)
        
        cnt = cnt + 1
    Next p
    
    ' Set frame scroll bar height
    scroll = (padTB * 4) + ((cnt - 1) * spacing)
    If scroll >= Me.VariablesFrame.Height Then
        Me.VariablesFrame.Scrollbars = fmScrollBarsVertical
        Me.VariablesFrame.ScrollHeight = scroll
    End If
End Sub

Private Function grabParamsValue() As Object
    Dim rtn As Object: Set rtn = CreateObject("Scripting.Dictionary")
    Dim i As Integer
    For i = 0 To (Me.VariablesFrame.Controls.Count - 1) Step 1
        Dim ctrl As MSForms.Control: Set ctrl = Me.VariablesFrame.Controls.item(i)
        If TypeOf ctrl Is MSForms.TextBox Then
            rtn.add ctrl.Name, ctrl.Value
        End If
    Next i
    
    Set grabParamsValue = rtn
End Function
