' --- Initialisation du UserForm ---
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim nom As String
    Dim numFiche As String
    
    Label2.Caption = Format(Date, "dddd dd mmmm yyyy")
    
    Set ws = ThisWorkbook.Sheets("database") 
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row 
    
   
    For Each cell In ws.Range("D10:D" & lastRow)
        If cell.Value <> "" Then
            nom = cell.Value 
            numFiche = ws.Cells(cell.Row, 5).Value 
            Me.ComboBox1.AddItem nom & " - " & numFiche 
        End If
    Next cell
End Sub

' --- Gestion du ComboBox ---
Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then '
        Dim ws As Worksheet
        Dim lastRow As Long
        Dim cell As Range
        Dim nomRecherche As String
        
        Set ws = ThisWorkbook.Sheets("database") 
        lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row 

        nomRecherche = Split(Me.ComboBox1.Text, " - ")(0) 
        
       
        For Each cell In ws.Range("D10:D" & lastRow)
            If cell.Value = nomRecherche Then
                
                Me.LabelDate.Caption = Format(cell.Offset(0, -1).Value, "dddd dd mmmm yyyy") ' Colonne C
                Me.LabelRef.Caption = cell.Offset(0, 1).Value ' Colonne E
                Me.LabelRef1.Caption = cell.Offset(0, 2).Value
                
                Me.MultiPage1.Value = 0 
                
                Exit For 
            End If
        Next cell
    End If
End Sub

' --- Auto-complétion du ComboBox ---
Private Sub ComboBox1_Change()
    Dim i As Integer
    Dim txt As String
    txt = Me.ComboBox1.Text
    
    
    Application.EnableEvents = False
    
    ' Vérifier si du texte est saisi
    If txt <> "" Then
        For i = 0 To Me.ComboBox1.ListCount - 1
            If LCase(Left(Me.ComboBox1.List(i), Len(txt))) = LCase(txt) Then
                Me.ComboBox1.Text = Me.ComboBox1.List(i)
                Me.ComboBox1.SelStart = Len(txt)
                Me.ComboBox1.SelLength = Len(Me.ComboBox1.Text) - Len(txt)
                Exit For
            End If
        Next i
    End If
    
    ' Réactive les événements
    Application.EnableEvents = True
End Sub

Private Sub btn_quitter_Click()
    End
End Sub

Sub Cons()
    Main.Show
End Sub


