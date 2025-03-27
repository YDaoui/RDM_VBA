VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Main 
   Caption         =   "Console"
   ClientHeight    =   11445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23295
   OleObjectBlob   =   "Main.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub btn_annuler_Click()
OptionButton_Existant.value = False
OptionButton_Nouveau.value = False
ComboBox1.Text = ""
Frame3.Visible = False
Frame2.Visible = False
TextBox_Searsh.Visible = True
End Sub

Private Sub btn_enregistrer_Click()
    ' Déclarer les variables
    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim valeur1 As String
    Dim valeur2 As String
    Dim valeur3 As String
    
    ' Définir la feuille "Database"
    Set ws = ThisWorkbook.Sheets("Database")
    
    ' Trouver la dernière ligne remplie dans le tableau (colonne C)
    derniereLigne = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row + 1
    
    ' Vérifier si la dernière ligne est en dessous de la ligne 10
    If derniereLigne < 10 Then
        derniereLigne = 10 ' Commencer à la ligne 10 si le tableau est vide
    End If
    
    ' Récupérer les valeurs des zones de texte
    valeur1 = Me.TextBox1.value
    valeur2 = Me.TextBox2.value
    valeur3 = Me.TextBox3.value
    
    ' Enregistrer les valeurs dans le tableau (colonnes C, D, E)
    ws.Cells(derniereLigne, 3).value = valeur1
    ws.Cells(derniereLigne, 4).value = valeur2
    ws.Cells(derniereLigne, 5).value = valeur3
    
    ' Afficher un message de confirmation
    MsgBox "Les données ont été enregistrées avec succès !", vbInformation
    
    ' Effacer les zones de texte après enregistrement
    Me.TextBox1.value = ""
    Me.TextBox2.value = ""
    Me.TextBox3.value = ""
End Sub

Private Sub btn_quitter_Click()
End
End Sub





Private Sub Combo_Action_Global_Change()
On Error GoTo ErrorHandler ' Activer la gestion des erreurs

    ' Vérifier si un élément est sélectionné dans ComboBox1
    If Me.Combo_Action_Global.ListIndex = -1 Then
        Exit Sub
    End If

    ' Récupérer la valeur sélectionnée dans ComboBox1
    Dim selectedValue As String
    selectedValue = Me.Combo_Action_Global.value

    ' Accéder à la feuille "Console"
    Dim wsConsole As Worksheet
    Set wsConsole = ThisWorkbook.Sheets("Console")

    ' Mettre à jour la cellule G35
    wsConsole.Range("H35").value = selectedValue
    wsConsole.Range("G35").value = "Suivie"
    wsConsole.Range("E35").value = Label_NumFiche.Caption

    Exit Sub

ErrorHandler:
    ' Gestion des erreurs
    MsgBox "Une erreur s'est produite : " & Err.Description, vbCritical, "Erreur"
End Sub


Private Sub Combo_Action_Nouveau_Change()




 Dim wsConsole As Worksheet
    Set wsConsole = ThisWorkbook.Sheets("Console")

    ' Mettre à jour la cellule G35
    wsConsole.Range("H35").value = selectedValue
    wsConsole.Range("G35").value = "Suivie"
    wsConsole.Range("E35").value = Label_NumFiche.Caption
    
    
    
    
    
End Sub

Private Sub Combo_Resume_existant_Change()
If Combo_Resume_existant.Text = "Rappel" Then

ComboBox11.AddItem Date
ComboBox11.AddItem Date + 1
ComboBox11.AddItem Date + 2
ComboBox11.AddItem Date + 3
ComboBox11.AddItem Date + 4
ComboBox11.AddItem Date + 5


Frame7.Visible = True

Else

Frame7.Visible = False
End If

End Sub

Private Sub Combo_Resume_Global_Change()

If Combo_Resume_Global.Text = "Rappel" Then

ComboBox7.AddItem Date
ComboBox7.AddItem Date + 1
ComboBox7.AddItem Date + 2
ComboBox7.AddItem Date + 3
ComboBox7.AddItem Date + 4
ComboBox7.AddItem Date + 5



Frame_Global_Rappel.Visible = True
Else

Frame_Global_Rappel.Visible = False
End If

On Error GoTo ErrorHandler ' Activer la gestion des erreurs

    ' Vérifier si un élément est sélectionné dans ComboBox1
    If Me.Combo_Resume_Global.ListIndex = -1 Then
        Exit Sub
    End If

    ' Récupérer la valeur sélectionnée dans ComboBox1
    Dim selectedValue As String
    selectedValue = Me.Combo_Resume_Global.value

    ' Accéder à la feuille "Console"
    Dim wsConsole As Worksheet
    Set wsConsole = ThisWorkbook.Sheets("Console")

    ' Mettre à jour la cellule G35
    wsConsole.Range("I35").value = selectedValue

    Exit Sub

ErrorHandler:
    ' Gestion des erreurs
    MsgBox "Une erreur s'est produite : " & Err.Description, vbCritical, "Erreur"


End Sub

Private Sub Combo_Resume_Nouveau_Change()
If Combo_Resume_Nouveau.Text = "Rappel" Then
ComboBox_DateRappel_Nouveau.AddItem Date
ComboBox_DateRappel_Nouveau.AddItem Date + 1
ComboBox_DateRappel_Nouveau.AddItem Date + 2
ComboBox_DateRappel_Nouveau.AddItem Date + 3
ComboBox_DateRappel_Nouveau.AddItem Date + 4
ComboBox_DateRappel_Nouveau.AddItem Date + 5
ComboBox_DateRappel_Nouveau.AddItem Date + 6
ComboBox_DateRappel_Nouveau.AddItem Date + 7

ComboBox_CrenauRappel_Nouveau.AddItem "9:00"
ComboBox_CrenauRappel_Nouveau.AddItem "9:30"
ComboBox_CrenauRappel_Nouveau.AddItem "10:00"
ComboBox_CrenauRappel_Nouveau.AddItem "10:30"




Frame8.Visible = True

Else
Frame8.Visible = False

End If

End Sub

Private Sub ComboBox6_Change()

On Error GoTo ErrorHandler ' Activer la gestion des erreurs

    ' Vérifier si un élément est sélectionné dans ComboBox1
    If Me.ComboBox6.ListIndex = -1 Then
        Exit Sub
    End If

    ' Récupérer la valeur sélectionnée dans ComboBox1
    Dim selectedValue As String
    selectedValue = Me.ComboBox6.value

    ' Accéder à la feuille "Console"
    Dim wsConsole As Worksheet
    Set wsConsole = ThisWorkbook.Sheets("Console")

    ' Mettre à jour la cellule G35
    wsConsole.Range("K35").value = selectedValue

    Exit Sub

ErrorHandler:
    ' Gestion des erreurs
    MsgBox "Une erreur s'est produite : " & Err.Description, vbCritical, "Erreur"
End Sub

Private Sub ComboBox7_Change()



On Error GoTo ErrorHandler ' Activer la gestion des erreurs

    ' Vérifier si un élément est sélectionné dans ComboBox1
    If Me.ComboBox7.ListIndex = -1 Then
        Exit Sub
    End If

    ' Récupérer la valeur sélectionnée dans ComboBox1
    Dim selectedValue As String
    selectedValue = Me.ComboBox7.value
    
    ComboBox6.AddItem "9:00"
   ComboBox6.AddItem "9:30"
   ComboBox6.AddItem "10:00"
   ComboBox6.AddItem "10:30"
   ComboBox6.AddItem "11:00"
   ComboBox6.AddItem "11:30"
   ComboBox6.AddItem "14:00"
   ComboBox6.AddItem "14:30"
   ComboBox6.AddItem "15:00"
   ComboBox6.AddItem "15:30"
   ComboBox6.AddItem "16:00"
   ComboBox6.AddItem "16:30"


    ' Accéder à la feuille "Console"
    Dim wsConsole As Worksheet
    Set wsConsole = ThisWorkbook.Sheets("Console")

    ' Mettre à jour la cellule G35
    wsConsole.Range("J35").value = selectedValue

    Exit Sub
    
ErrorHandler:
    ' Gestion des erreurs
    MsgBox "Une erreur s'est produite : " & Err.Description, vbCritical, "Erreur"
End Sub

Private Sub ComboDate_Rappel_Global_Change()
    UpdateTabCaption

    Dim wsDatabase As Worksheet
    Dim lastRowDatabase As Long
    Dim i As Long
    Dim selectedDate As Date
    Dim numFiche As String
    Dim maxDate As Date, maxTime As Date, maxRow As Long

    ' Définir la feuille
    On Error Resume Next
    Set wsDatabase = ThisWorkbook.Sheets("Actions_Fiches")
    On Error GoTo 0

    ' Vérifier si la feuille existe
    If wsDatabase Is Nothing Then
        MsgBox "La feuille 'Actions_Fiches' n'existe pas !", vbCritical, "Erreur"
        Exit Sub
    End If

    ' Vérifier si ComboDate_Rappel_Global contient une date valide
    If IsDate(ComboDate_Rappel_Global.value) Then
        selectedDate = CDate(ComboDate_Rappel_Global.value)
    Else
        MsgBox "Sélection invalide dans ComboDate_Rappel_Global.", vbExclamation, "Erreur"
        Exit Sub
    End If

    ' Trouver la dernière ligne utilisée dans la colonne C
    lastRowDatabase = wsDatabase.Cells(wsDatabase.Rows.Count, "C").End(xlUp).Row

    ' Effacer les anciens éléments de la ListBox
    Me.ListBox_Global_Rappels.Clear
    Me.ListBox_Global_Rappels.ColumnCount = 8 ' Nombre de colonnes
    Me.ListBox_Global_Rappels.ColumnWidths = "80;80;80;80;80;80;80;80"
    Me.ListBox_Global_Rappels.BackColor = RGB(200, 200, 255)
    Me.ListBox_Global_Rappels.Font.Bold = True

    ' Ajouter les en-têtes à ListBox_Global_Rappels
    Me.ListBox_Global_Rappels.AddItem
    Me.ListBox_Global_Rappels.List(0, 0) = "Date Rappel"
    Me.ListBox_Global_Rappels.List(0, 1) = "Creneau"
    Me.ListBox_Global_Rappels.List(0, 2) = "Assu"
    Me.ListBox_Global_Rappels.List(0, 3) = "Action"
    Me.ListBox_Global_Rappels.List(0, 4) = "Commentaire"
    Me.ListBox_Global_Rappels.List(0, 5) = "Motif"
    Me.ListBox_Global_Rappels.List(0, 6) = "Résumé"
    Me.ListBox_Global_Rappels.List(0, 7) = "NumFiche"

    ' Parcourir la feuille "Actions_Fiches" pour filtrer les rappels
    For i = 10 To lastRowDatabase
        ' Vérifier les conditions : statut non "Clôt", type "Rappel", et statut "Prévu"
        If wsDatabase.Cells(i, "K").value <> "Clôt" And _
           wsDatabase.Cells(i, "L").value = "Rappel" And _
           wsDatabase.Cells(i, "P").value = "Prévu" Then

            ' Vérifier si la date dans la colonne N correspond à la date sélectionnée
            If IsDate(wsDatabase.Cells(i, "N").value) Then
                If wsDatabase.Cells(i, "N").value >= selectedDate Then
                    ' Récupérer le NumFiche correspondant
                    numFiche = wsDatabase.Cells(i, "E").value

                    ' Initialiser les variables pour stocker la date et l'heure maximales
                    maxDate = DateSerial(1900, 1, 1)
                    maxTime = TimeSerial(0, 0, 0)
                    maxRow = -1

                    ' Parcourir à nouveau pour trouver le rappel le plus récent pour ce NumFiche
                    For j = 10 To lastRowDatabase
                        If wsDatabase.Cells(j, "E").value = numFiche And _
                           wsDatabase.Cells(j, "K").value <> "Clôt" And _
                           wsDatabase.Cells(j, "L").value = "Rappel" And _
                           wsDatabase.Cells(j, "P").value = "Prévu" Then

                            ' Vérifier si la date et l'heure sont plus récentes
                            If wsDatabase.Cells(j, "N").value > maxDate Or _
                               (wsDatabase.Cells(j, "N").value = maxDate And wsDatabase.Cells(j, "O").value > maxTime) Then
                                maxDate = wsDatabase.Cells(j, "N").value
                                maxTime = wsDatabase.Cells(j, "O").value
                                maxRow = j ' Enregistrer la ligne du rappel le plus récent
                            End If
                        End If
                    Next j

                    ' Si un rappel a été trouvé, ajouter les informations à la ListBox
                    If maxRow <> -1 Then
                        Me.ListBox_Global_Rappels.AddItem
                        Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 0) = Format(wsDatabase.Cells(maxRow, "N").value, "dddd dd/mm")
                        Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 1) = Format(wsDatabase.Cells(maxRow, "O").value, "hh:mm")
                        Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 2) = wsDatabase.Cells(maxRow, "C").value
                        Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 3) = wsDatabase.Cells(maxRow, "I").value
                        Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 4) = wsDatabase.Cells(maxRow, "M").value
                        Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 5) = wsDatabase.Cells(maxRow, "H").value
                        Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 6) = wsDatabase.Cells(maxRow, "P").value
                        Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 7) = numFiche
                    End If
                End If
            End If
        End If
    Next i

    ' Vérifier si la ListBox est vide après le filtrage
    If Me.ListBox_Global_Rappels.ListCount = 0 Then
        MsgBox "Aucun rappel trouvé pour cette date.", vbInformation, "Info"
        Label20.Visible = False
    Else
        Label20.Visible = True
    End If

End Sub

Private Sub CommandButton1_Click()

    ' Vérifie si l'utilisateur appuie sur la touche Entrée
    
        Dim inputText As String
        Dim nom As String
        
        ' Récupérer le texte saisi dans ComboBox1
        inputText = Me.TextBox_Searsh.Text
        
        ' Extraire le nom après le tiret
        If InStr(inputText, " - ") > 0 Then
            nom = Trim(Split(inputText, " - ")(1)) ' Prendre la partie après le tiret
        Else
            nom = inputText ' Si le format n'est pas respecté, utiliser tout le texte
        End If
        
        ' Afficher le nom dans TextBox2
        Me.ComboBox1.Text = nom
    

End Sub

Private Sub CommandButton2_Click()
Dim wsConsole As Worksheet
    Set wsConsole = ThisWorkbook.Sheets("Console")
    
    
    wsConsole.Range("D1").value = Label_NumFiche.Caption
    
    wsConsole.Range("D1").Copy
    
    
    
End Sub



Private Sub CommandButton5_Click()

End Sub

Private Sub CommandButton7_Click()
Dim wsConsole As Worksheet
    Set wsConsole = ThisWorkbook.Sheets("Console")
    
    
    wsConsole.Range("E1").value = Label_Assu.Caption
    
    wsConsole.Range("E1").Copy
End Sub

Private Sub CommandButton8_Click()
TextBox_Searsh.Text = ""
ComboBox1.Text = ""
OptionButton_Global.value = True
End Sub

Private Sub CommandButton9_Click()
 Dim wsConsole As Worksheet, ws_Fiches As Worksheet, wsActions_Fiches As Worksheet




 Set wsConsole = ThisWorkbook.Sheets("Console")
 Set ws_Fiches = ThisWorkbook.Sheets("Fiches")
 Set wsActions_Fiches = ThisWorkbook.Sheets("Actions_Fiches")




If txtRef.Text <> "$" Then


'Enregistrement de la fiche

wsConsole.Range("E28").value = Me.txtRef.Text 'NumFiche
wsConsole.Range("F28").value = Now 'Date Fiche

wsConsole.Range("G28").value = Me.TextBox_Searsh.value 'Ref


wsConsole.Range("H28").value = Combo_Source_Nouveau.Text 'Source

wsConsole.Range("I28").value = Combo_Motif_Nouveau 'Motif

    If CheckBox2.value = False Then
        wsConsole.Range("J28").value = "Normale"
    Else
        wsConsole.Range("J28").value = "Urgent"
    End If
wsConsole.Range("L28").value = TextBox4.Text ' Comentaire


wsConsole.Range("E28:L28").Copy


For i = 10 To ws_Fiches.Cells(ws_Fiches.Rows.Count, "H").End(xlUp).Row
    
   
    If Me.txtRef.Text = ws_Fiches.Cells(i, "H").value Then
        MsgBox "Error: La référence existe déjà."
    Else
       
        
        
      
        lastRow = ws_Fiches.Cells(ws_Fiches.Rows.Count, "H").End(xlUp).Row + 1
        
       
        ws_Fiches.Cells(lastRow, "h").PasteSpecial Paste:=xlPasteValues
        
      
        Exit For
    End If
Next i

' Désactiver le mode couper-copier
Application.CutCopyMode = False

'Enregistrement d'Action sur  la fiche


'Num Fiche C
wsConsole.Range("C31").value = Me.txtRef.value 'NumFiche
'Date Action D
wsConsole.Range("D31").value = Now 'Date Fiche
'Source : E
wsConsole.Range("E31").value = Combo_Source_Nouveau.Text 'Source
'Motif  Action   F
wsConsole.Range("F31").value = Combo_Motif_Nouveau 'Motif
'Action_Fiche   G
wsConsole.Range("G31").value = Combo_Action_Nouveau 'Action
'Priorit H
If CheckBox2.value = False Then
        wsConsole.Range("H31").value = "Normale"
    Else
        wsConsole.Range("H31").value = "Urgent"
    End If
'Statut Action  I
wsConsole.Range("I31").value = Combo_Resume_Nouveau
'Resumé Action Fiche  J
wsConsole.Range("J31").value = Combo_Resume_Nouveau
'Commentaire Action Fiche  K
wsConsole.Range("K31").value = TextBox8.Text ' Comentaire
'Date Rappel  L
wsConsole.Range("L31").value = ComboBox_DateRappel_Nouveau.Text
'Creneau Rappel  M
wsConsole.Range("M31").value = ComboBox_CrenauRappel_Nouveau.Text
'Statut Rappel  N
wsConsole.Range("N31").value = "Prévu"



wsConsole.Range("C31:N31").Copy


For i = 10 To wsActions_Fiches.Cells(wsActions_Fiches.Rows.Count, "H").End(xlUp).Row + 1
    
   
   
       
        
        
      
        'lastRow = wsActions_Fiches.Cells(ws_Fiches.Rows.Count, "H").End(xlUp).Row + 1
        
       
        wsActions_Fiches.Cells(lastRow, "F").PasteSpecial Paste:=xlPasteValues
        
      
        Exit For

Next i



Else
MsgBox "Attention"

End If


End Sub

Private Sub List_Histo_Click()

Dim numFiche As String

    If List_Histo.ListIndex <> -1 Then
    
    'Label32.Visible = False
       
        LabelRef.Caption = List_Histo.List(List_Histo.ListIndex, 6)
        'numFiche = ListBox_Global.List(ListBox_Global.ListIndex, 7)
        LabelRef.Visible = True
       ' Label_NumFiche.Caption = numFiche
       ' Label_NumFiche.Visible = True
        Frame6.Visible = True
    Else
        LabelRef.Caption = "Aucune sélection"
        LabelRef.Visible = False
    End If
    

    'LabelRef.Visible = True
    
End Sub


Private Sub ListBox_Global_Click()
    Dim wsDatabase As Worksheet
    Dim lastRowDatabase As Long
    Dim selectedNumFiche As String
    Dim i As Long

    ' Trouver la feuille "database"
    On Error Resume Next
    Set wsDatabase = ThisWorkbook.Sheets("Actions_Fiches")
    On Error GoTo 0

    

    ' Déterminer la dernière ligne de la feuille "database"
    lastRowDatabase = wsDatabase.Cells(wsDatabase.Rows.Count, "E").End(xlUp).Row


    ' Vérifier si un élément est sélectionné dans ListBox_Global
    

    ' Récupérer le NumFiche de l'élément sélectionné
    selectedNumFiche = Me.ListBox_Global.List(Me.ListBox_Global.ListIndex, 7)
     Label__Global_Liste_Commentaire.Caption = ListBox_Global.List(ListBox_Global.ListIndex, 4)
    
    
    ' Récupérer le commentaire de la ligne sélectionnée dans ListBox_Global
    Dim selectedCommentaire As String
    selectedCommentaire = Me.ListBox_Global.List(Me.ListBox_Global.ListIndex, 4) ' Commentaire est dans la colonne 4

    ' Afficher le commentaire dans le label
    Label__Global_Liste_Commentaire.Caption = selectedCommentaire
    Label__Global_Liste_Commentaire.Visible = True

    ' Effacer la ListBox_Global_Details pour afficher les nouvelles informations
    Me.ListBox_Global_Details.Clear
    Me.ListBox_Global_Details.ColumnCount = 8
    Me.ListBox_Global_Details.ColumnWidths = "80;80;80;80;80;80"
    Me.ListBox_Global_Details.BackColor = RGB(200, 200, 255)
    Me.ListBox_Global_Details.Font.Bold = True

    ' Ajouter les en-têtes à ListBox_Global_Details
    Me.ListBox_Global_Details.AddItem
    Me.ListBox_Global_Details.List(0, 0) = "Date"
    Me.ListBox_Global_Details.List(0, 1) = "Heure"
    Me.ListBox_Global_Details.List(0, 2) = "Assu"
    Me.ListBox_Global_Details.List(0, 3) = "Action"
    Me.ListBox_Global_Details.List(0, 4) = "Commentaire"
    Me.ListBox_Global_Details.List(0, 5) = "Motif"
    Me.ListBox_Global_Details.List(0, 6) = "Résumé"
    Me.ListBox_Global_Details.List(0, 7) = "NumFiche"

    ' Parcourir la base de données pour trouver les enregistrements correspondants au NumFiche sélectionné
    For i = 10 To lastRowDatabase
        If wsDatabase.Cells(i, "E").value = selectedNumFiche Then ' Vérifier si le NumFiche correspond
            ' Ajouter l'enregistrement à ListBox_Global_Details
            Me.ListBox_Global_Details.AddItem
            Me.ListBox_Global_Details.List(Me.ListBox_Global_Details.ListCount - 1, 0) = Format(wsDatabase.Cells(i, "B").value, "dddd dd/mm") ' Date Rappel (colonne K)
            Me.ListBox_Global_Details.List(Me.ListBox_Global_Details.ListCount - 1, 1) = Format(wsDatabase.Cells(i, "A").value, "hh:mm") ' Créneau (colonne L)
            Me.ListBox_Global_Details.List(Me.ListBox_Global_Details.ListCount - 1, 2) = wsDatabase.Cells(i, "D").value ' Assu (colonne D)
            Me.ListBox_Global_Details.List(Me.ListBox_Global_Details.ListCount - 1, 3) = wsDatabase.Cells(i, "I").value ' Action (colonne I)
            Me.ListBox_Global_Details.List(Me.ListBox_Global_Details.ListCount - 1, 4) = wsDatabase.Cells(i, "M").value ' Commentaire (colonne M)
            Me.ListBox_Global_Details.List(Me.ListBox_Global_Details.ListCount - 1, 5) = wsDatabase.Cells(i, "H").value ' Motif (colonne H)
            Me.ListBox_Global_Details.List(Me.ListBox_Global_Details.ListCount - 1, 6) = wsDatabase.Cells(i, "J").value ' Résumé (colonne J)
            Me.ListBox_Global_Details.List(Me.ListBox_Global_Details.ListCount - 1, 7) = wsDatabase.Cells(i, "C").value ' NumFiche (colonne C)
        End If
    Next i

    ' Mettre à jour les labels avec les informations de l'élément sélectionné
    If Me.ListBox_Global_Details.ListCount > 0 Then
    
    For i = 10 To wsDatabase.Cells(wsDatabase.Rows.Count, "E").End(xlUp).Row
        
     If wsDatabase.Cells(i, "E").value = selectedNumFiche Then
  
          Label_Assu.Caption = wsDatabase.Cells(i, "C").value
          Label_Date_Fiche.Caption = Format(wsDatabase.Cells(i, "F").value, "dddd dd mmmm  hh:mm:ss ")
          
     
        Label_NumFiche.Caption = selectedNumFiche ' NumFiche
        Label_NumFiche.Visible = True
        Frame5.Visible = True
        ListBox_Global_Details.Visible = True
       
      Label34.Visible = True
        Label35.Visible = False

       
       
        Label__Global_Liste_Details_Commentaire.Visible = False
    Else
        'Label__Global_Liste_Commentaire.Caption = "Aucune information trouvée"
    End If
    Next i
    
    End If
    
End Sub
Sub Assu_Info()

 On Error Resume Next
    Set wsDatabase = ThisWorkbook.Sheets("database")
    On Error GoTo 0

    ' Vérifier si la feuille "database" existe
    If wsDatabase Is Nothing Then
        MsgBox "La feuille 'database' n'existe pas !", vbCritical, "Erreur"
        Exit Sub
    End If
    
For i = 10 To wsDatabase.Cells(wsDatabase.Rows.Count, "C").End(xlUp).Row
        
     If wsDatabase.Cells(i, "C").value = selectedNumFiche Then
     MsgBox "lkjqshfqlkhf"
          Label_Assu.Caption = wsDatabase.Cells(i, "D").value
          Label_Date_Fiche.Caption = Format(wsDatabase.Cells(i, "E").value, "dddd dd mmmm  hh:mm:ss ")
     
        Label_NumFiche.Caption = selectedNumFiche ' NumFiche
        Label_NumFiche.Visible = True
        Frame5.Visible = True
        ListBox_Global_Details.Visible = True
    Else
        Label__Global_Liste_Commentaire.Caption = "Aucune information trouvée"
    End If
    Next i
    
End Sub


Private Sub ListBox_Global_Details_Click()
Dim numFiche As String



    If ListBox_Global_Details.ListIndex <> -1 Then
       'Label34.Visible = False
Label35.Visible = True
        Label__Global_Liste_Details_Commentaire.Caption = ListBox_Global_Details.List(ListBox_Global_Details.ListIndex, 4)
        'numFiche = ListBox_Global_Details.List(ListBox_Global_Details.ListIndex, 7)
        'Label_NumFiche.Caption = numFiche
        'Label_NumFiche.Visible = True
        Frame5.Visible = True
    Else
        Label__Global_Liste_Details_Commentaire.Caption = "Aucune sélection"
    End If
    
    
    Label__Global_Liste_Details_Commentaire.Visible = True
    Frame_Action_Global.Visible = True
End Sub

Private Sub ListBox_Global_Rappels_Click()
  Dim numFiche As String


Set wsDatabase = ThisWorkbook.Sheets("database")
    On Error GoTo 0

    ' Vérifier si la feuille "database" existe
    If wsDatabase Is Nothing Then
        MsgBox "La feuille 'database' n'existe pas !", vbCritical, "Erreur"
        Exit Sub
    End If
    
    
    

    If ListBox_Global_Rappels.ListIndex <> -1 Then
       
        Label_Global_Rappel.Caption = ListBox_Global_Rappels.List(ListBox_Global_Rappels.ListIndex, 4)
        
        numFiche = ListBox_Global_Rappels.List(ListBox_Global_Rappels.ListIndex, 7)
        
       For i = 10 To wsDatabase.Cells(wsDatabase.Rows.Count, "C").End(xlUp).Row
        
     If wsDatabase.Cells(i, "C").value = numFiche Then
  
          Label_Assu.Caption = wsDatabase.Cells(i, "D").value
          Label_Date_Fiche.Caption = Format(wsDatabase.Cells(i, "E").value, "dddd dd mmmm  hh:mm:ss ")
     
        Label_NumFiche.Caption = numFiche ' NumFiche
        Label_NumFiche.Visible = True
        Frame5.Visible = True
        ListBox_Global_Details.Visible = True
    Else
        Label__Global_Liste_Commentaire.Caption = "Aucune information trouvée"
    End If
    Next i
    
    End If
    
    
    Label_Global_Rappel.Visible = True
    Frame_Action_Global.Visible = True

End Sub

Private Sub ListBox_Rappel_Click()

Dim numFiche As String



    If ListBox_Rappel.ListIndex <> -1 Then
        '
        
        LabelRef.Caption = ListBox_Rappel.List(ListBox_Rappel.ListIndex, 6)
        'numFiche = ListBox_Rappel.List(ListBox_Rappel.ListIndex, 7)
        'Label_NumFiche.Caption = numFiche
        'Label_NumFiche.Visible = True
        Frame6.Visible = True
    Else
    
        LabelRef.Caption = "Aucune sélection"
    End If
    
    
    LabelRef.Visible = True
End Sub

Private Sub ListBox_Rappel_Histo_Click()
Dim numFiche As String



    If ListBox_Rappel_Histo.ListIndex <> -1 Then
        '
        
        LabelRef.Caption = ListBox_Rappel_Histo.List(ListBox_Rappel_Histo.ListIndex, 6)
        'numFiche = ListBox_Rappel.List(ListBox_Rappel.ListIndex, 7)
        'Label_NumFiche.Caption = numFiche
        'Label_NumFiche.Visible = True
        Frame6.Visible = True
    Else
    
        LabelRef.Caption = "Aucune sélection"
    End If
    
    
    LabelRef.Visible = True
End Sub

Private Sub MultiPage_Global_Change()

End Sub

Private Sub MultiPage1_Change()
If MultiPage1.value = 1 Then
Label19.Visible = True
ComboDate_Rappel_Global.Visible = True

'MsgBox "L'événement MultiPage1_Change est déclenché. Onglet sélectionné : " & MultiPage1.value

Else
Label19.Visible = False
ComboDate_Rappel_Global.Visible = False
End If
End Sub
    
'End Sub
Private Sub MultiPage2_Change()
    MsgBox "L'événement MultiPage1_Change est déclenché. yyyyyyyyyyyyyyyOnglet sélectionné : " & MultiPage1.value
End Sub

Private Sub OptionButton_Global_Click()
Frame_Stat.Visible = False
Frame3.Visible = False
Frame2.Visible = False
Frame4.Visible = True
Frame_Action_Global.Visible = False
Label__Global_Liste_Commentaire.Visible = False
Label__Global_Liste_Details_Commentaire.Visible = False

Label34.Visible = False
Label35.Visible = False


    'Select Case MultiPage1.value
       ' Case 0 ' Page 1 (Détails)
          '  Me.ComboDate_Rappel_Global.Visible = False
          '  Label19.Visible = False
            
    ' MsgBox "MultiPage1_Change déclenché ! Onglet sélectionné : " & MultiPage1.value
      '  Case 1 ' Page 2 (Rappels)
            'Me.ComboDate_Rappel_Global.Visible = True
          '  Label19.Visible = True
   ' End Select

Combo_Action_Global.AddItem "Appel"
Combo_Action_Global.AddItem "Mail"
Combo_Action_Global.AddItem "Proposition"
Combo_Action_Global.AddItem "Etude"





   Combo_Resume_Global.AddItem "Rappel"
   Combo_Resume_Global.AddItem "Suivie"
   Combo_Resume_Global.AddItem "Traité"
   Combo_Resume_Global.AddItem "Injoignable"
   
  
   
   
Dim wsDatabase As Worksheet
Dim cell As Range
Dim lastRowDatabase As Long
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")


On Error Resume Next
Set wsDatabase = ThisWorkbook.Sheets("database")
On Error GoTo 0

If wsDatabase Is Nothing Then
    MsgBox "La feuille 'database' n'existe pas !", vbCritical, "Erreur"
    Exit Sub
End If


lastRowDatabase = wsDatabase.Cells(wsDatabase.Rows.Count, "K").End(xlUp).Row



ComboDate_Rappel_Global.Clear



For Each cell In wsDatabase.Range("K11:K" & lastRowDatabase)
    If Not dict.Exists(cell.value) And cell.value <> "" Then
        dict.Add cell.value, Nothing
        ComboDate_Rappel_Global.AddItem cell.value
    End If
Next cell

ListBox_Global_Details.Visible = False




Call Global_Stat

End Sub

Private Sub OptionButton_Nouveau_Click()
Frame3.Visible = True
Frame2.Visible = False
Frame4.Visible = False
TextBox_Searsh.Visible = True
ComboBox1.Text = ""

End Sub

Private Sub OptionButton_Existant_Click()
Frame2.Visible = True
Frame3.Visible = False
Frame4.Visible = False
Frame_Stat.Visible = False
End Sub






Private Sub OptionButton_stat_Click()

Frame_Stat.Visible = True
Frame2.Visible = False
Frame_Action_Global.Visible = False



End Sub

Private Sub txtRef_Change()
Sheets("Console").Range("E7").value = Me.txtRef.value
End Sub
Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CommandButton1.BackColor = RGB(0, 120, 215) ' Bleu clair au survol
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CommandButton1.BackColor = RGB(240, 240, 240)
     CommandButton2.BackColor = RGB(240, 240, 240)
     CommandButton7.BackColor = RGB(240, 240, 240) ' Retour à la couleur d'origine
End Sub
Private Sub CommandButton2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CommandButton2.BackColor = RGB(0, 120, 215) ' Bleu clair au survol
End Sub



Private Sub CommandButton7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CommandButton7.BackColor = RGB(0, 120, 215) ' Bleu clair au survol
End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Application.WindowState = xlNormal
End Sub
Sub LoadGrapik(SChart As String, ImgCtrl As Control, Optional FileIndex As Integer = 1)
    Dim NmChart As Chart
    Dim NmFile As String
    Dim UniqueID As String
    
    ' Générer un ID unique pour le fichier temporaire
    UniqueID = "Chart_" & FileIndex
    NmFile = ThisWorkbook.Path & "\" & UniqueID & ".gif"
    
    Set NmChart = ThisWorkbook.Sheets("Graphs").ChartObjects(SChart).Chart
    
    ' Exporter le graphique seulement si le fichier n'existe pas déjà
    If Dir(NmFile) = "" Then
        NmChart.Export NmFile, "GIF"
    End If
    
    On Error Resume Next
    Set ImgCtrl.Picture = LoadPicture(NmFile)
    If Err.Number <> 0 Then
        MsgBox "Impossible de charger l'image dans le contrôle.", vbExclamation
    End If
    On Error GoTo 0
End Sub




Private Sub UserForm_Initialize()
  Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim nom As String
    Dim numFiche As String
    
Call LoadGrapik("Graphique_Bar", Me.Image2, 1)
    Call LoadGrapik("Graphique_Nbr_Statut", Me.Image4, 2)
    Call LoadGrapik("Graphique_Total", Me.Image5, 3)


' Colorier le fond de MultiPage1 en blanc
MultiPage1.BackColor = RGB(255, 255, 255) ' Blanc

' Colorier le fond de MultiPage_Global en gris clair
MultiPage_Global.BackColor = RGB(192, 192, 192) ' Gris clair


   
    With Me
        .StartUpPosition = 0 ' Position manuelle
        .Left = (Application.Left + (Application.Width - .Width) / 2)
        .Top = (Application.Top + (Application.Height - .Height) / 2)
    End With


    
    
    
    
    
    
    
    
    With Me.TextBox_Searsh
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(200, 200, 200) ' Bordure grise
    End With
    With Me.TextBox6
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(200, 200, 200) ' Bordure grise
    End With
  
    
    
   Label2.Caption = Format(Date, "dddd dd mmmm yyyy")
   
   
   Combo_action_Existant.AddItem "Appel"
   Combo_action_Existant.AddItem "Mail"
   Combo_action_Existant.AddItem "Proposition"
   
   Combo_Action_Nouveau.AddItem "Appel"
   Combo_Action_Nouveau.AddItem "Mail"
   Combo_Action_Nouveau.AddItem "Proposition"
   
   
   
   Combo_Resume_Nouveau.AddItem "Clôt"
   Combo_Resume_Nouveau.AddItem "Rappel"
   Combo_Resume_Nouveau.AddItem "Suivie"
   Combo_Resume_Nouveau.AddItem "Traité"
   Combo_Resume_Nouveau.AddItem "Injoignable"
   
   
   Combo_Resume_existant.AddItem "Clôt"
   Combo_Resume_existant.AddItem "Rappel"
   Combo_Resume_existant.AddItem "Suivie"
   Combo_Resume_existant.AddItem "Traité"
   Combo_Resume_existant.AddItem "Injoignable"
   
   Combo_Source_Nouveau.AddItem "AMM"
   Combo_Source_Nouveau.AddItem "Assuré -Mail"
   Combo_Source_Nouveau.AddItem "Appel reçu"
   Combo_Source_Nouveau.AddItem "STANDARD -Murielle"
   Combo_Source_Nouveau.AddItem "STANDARD -Maeva"
   Combo_Source_Nouveau.AddItem "STANDARD -Anais"
   Combo_Source_Nouveau.AddItem "Compagne"
   
   Combo_Motif_Nouveau.AddItem "Explication"
   Combo_Motif_Nouveau.AddItem "Reclamation"
   Combo_Motif_Nouveau.AddItem "Point à refaire"
   
   Combo_Creneau.AddItem "9:00"
   Combo_Creneau.AddItem "9:30"
   Combo_Creneau.AddItem "10:00"
   Combo_Creneau.AddItem "10:30"
   Combo_Creneau.AddItem "11:00"
   Combo_Creneau.AddItem "11:30"
   Combo_Creneau.AddItem "14:00"
   Combo_Creneau.AddItem "14:30"
   Combo_Creneau.AddItem "15:00"
   Combo_Creneau.AddItem "15:30"
   Combo_Creneau.AddItem "16:00"
   Combo_Creneau.AddItem "16:30"
   
       
Set ws = ThisWorkbook.Sheets("Fiches")
lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row ' Trouve la dernière ligne dans la colonne F

For Each cell In ws.Range("F11:F" & lastRow)
    If cell.value <> "" Then
        ' Récupérer les valeurs en utilisant les lettres de colonnes
        nom = cell.value ' Colonne F
        numFiche = ws.Range("D" & cell.Row).value ' Colonne D
        datefiche = ws.Range("E" & cell.Row).value ' Colonne E
        
        ' Ajouter l'élément à la ComboBox
        Me.ComboBox1.AddItem nom & " - " & numFiche & " - " & datefiche
    End If
Next cell
  UpdateTabCaption
ListBox_Global_Details.Visible = False

    OptionButton_Global.value = True
    Application.WindowState = xlMinimized
End Sub

Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Vérifie si la touche "Entrée" est pressée
    If KeyCode = vbKeyReturn Then
        Dim wsFiches As Worksheet, wsConsole As Worksheet, wsFiches_Action As Worksheet
        Dim lastRowDatabase As Long, lastRowFichesAction As Long
        Dim cell As Range
        Dim nomRecherche As String, numFiche As String
        Dim found As Boolean
        Dim i As Long
        
        ' Initialisation des feuilles
        Set wsFiches = ThisWorkbook.Sheets("Fiches")
        Set wsFiches_Action = ThisWorkbook.Sheets("Actions_Fiches")
        Set wsConsole = ThisWorkbook.Sheets("Console")
        
        ' Trouver la dernière ligne dans les colonnes pertinentes
        lastRowDatabase = wsFiches.Cells(wsFiches.Rows.Count, "F").End(xlUp).Row
        lastRowFichesAction = wsFiches_Action.Cells(wsFiches_Action.Rows.Count, "E").End(xlUp).Row
        
        ' Extraire le nom recherché à partir de la ComboBox
        nomRecherche = Split(Me.ComboBox1.Text, " - ")(0)
        
        ' Initialisation de la variable found
        found = False
        
        ' Parcourir la colonne F de la feuille "Fiches"
        For Each cell In wsFiches.Range("F10:F" & lastRowDatabase)
            If cell.value = nomRecherche Then
                ' Récupérer les valeurs des colonnes correspondantes
                numFiche = wsFiches.Range("H" & cell.Row).value ' Colonne H
                Me.LabelDate.Caption = Format(wsFiches.Range("I" & cell.Row).value, "dddd dd mmmm yyyy  HH:MM") ' Colonne I
                Me.LabelRef.Caption = wsFiches.Range("K" & cell.Row).value ' Colonne K
                Me.LabelRef1.Caption = wsFiches.Range("N" & cell.Row).value ' Colonne N
                
                ' Mettre à jour les labels supplémentaires
                Me.Label_Assu.Caption = wsFiches.Range("F" & cell.Row).value ' Colonne F
                Me.Label_Date_Fiche.Caption = Format(wsFiches.Range("I" & cell.Row).value, "dddd dd mmmm  hh:mm:ss") ' Colonne I
                Label_Commentaire_Existant.Caption = wsFiches.Range("O" & cell.Row).value
                Me.Label_NumFiche.Caption = numFiche ' NumFiche
                Me.Label_NumFiche.Visible = True
                Me.Frame5.Visible = True
                Me.ListBox_Global_Details.Visible = True
                Me.Label__Global_Liste_Details_Commentaire.Visible = False
                
                ' Vider et configurer la ListBox pour l'historique
                Me.List_Histo.Clear
                Me.List_Histo.ColumnCount = 7
                Me.List_Histo.ColumnWidths = "100;60;100;100;100;100"
                Me.List_Histo.BackColor = RGB(200, 200, 255)
                Me.List_Histo.Font.Bold = True
                
                ' Ajouter les en-têtes de colonnes
                Me.List_Histo.AddItem
                Me.List_Histo.List(0, 0) = "Date"
                Me.List_Histo.List(0, 1) = "Heure"
                Me.List_Histo.List(0, 2) = "Source"
                Me.List_Histo.List(0, 3) = "Motif"
                Me.List_Histo.List(0, 4) = "Action"
                Me.List_Histo.List(0, 5) = "Statut"
                Me.List_Histo.List(0, 6) = "Commentaire"
                
                ' Remplir la ListBox avec les données de "Actions_Fiches"
                For i = 10 To lastRowFichesAction
                    If wsFiches_Action.Cells(i, "E").value = numFiche Then
                        Me.List_Histo.AddItem
                        Me.List_Histo.List(Me.List_Histo.ListCount - 1, 0) = Format(wsFiches_Action.Cells(i, "B").value, "dddd dd mmmm")
                        Me.List_Histo.List(Me.List_Histo.ListCount - 1, 1) = wsFiches_Action.Cells(i, "A").value
                        Me.List_Histo.List(Me.List_Histo.ListCount - 1, 2) = wsFiches_Action.Cells(i, "G").value
                        Me.List_Histo.List(Me.List_Histo.ListCount - 1, 3) = wsFiches_Action.Cells(i, "H").value
                        Me.List_Histo.List(Me.List_Histo.ListCount - 1, 4) = wsFiches_Action.Cells(i, "I").value
                        Me.List_Histo.List(Me.List_Histo.ListCount - 1, 5) = wsFiches_Action.Cells(i, "K").value
                        Me.List_Histo.List(Me.List_Histo.ListCount - 1, 6) = wsFiches_Action.Cells(i, "M").value
                    End If
                Next i
                
                ' Vider et configurer la ListBox pour les rappels
                Me.ListBox_Rappel.Clear
                Me.ListBox_Rappel.ColumnCount = 6
                Me.ListBox_Rappel.ColumnWidths = "100;80;80;80;90;80"
                Me.ListBox_Rappel.BackColor = RGB(200, 200, 255)
                Me.ListBox_Rappel.Font.Bold = True
                
                ' Ajouter les en-têtes de colonnes
                Me.ListBox_Rappel.AddItem
                Me.ListBox_Rappel.List(0, 0) = "Date Rappel"
                Me.ListBox_Rappel.List(0, 1) = "Creneau"
                Me.ListBox_Rappel.List(0, 2) = "Motif"
                Me.ListBox_Rappel.List(0, 3) = "Action"
                Me.ListBox_Rappel.List(0, 4) = "Statut"
                Me.ListBox_Rappel.List(0, 5) = "Statut Rappel"
                Me.ListBox_Rappel.List(0, 6) = "Commentaire"
                
                ' Remplir la ListBox avec les rappels
                For i = 10 To lastRowFichesAction
                    If wsFiches_Action.Cells(i, "E").value = numFiche And wsFiches_Action.Cells(i, "L").value = "Rappel" Then
                        Me.ListBox_Rappel.AddItem
                        Me.ListBox_Rappel.List(Me.ListBox_Rappel.ListCount - 1, 0) = Format(wsFiches_Action.Cells(i, "N").value, "dddd dd/mmmm")
                        Me.ListBox_Rappel.List(Me.ListBox_Rappel.ListCount - 1, 1) = Format(wsFiches_Action.Cells(i, "O").value, "hh:mm")
                        Me.ListBox_Rappel.List(Me.ListBox_Rappel.ListCount - 1, 2) = wsFiches_Action.Cells(i, "H").value 'Motif
                        Me.ListBox_Rappel.List(Me.ListBox_Rappel.ListCount - 1, 3) = wsFiches_Action.Cells(i, "K").value 'Action
                        Me.ListBox_Rappel.List(Me.ListBox_Rappel.ListCount - 1, 4) = wsFiches_Action.Cells(i, "L").value 'Statut
                        Me.ListBox_Rappel.List(Me.ListBox_Rappel.ListCount - 1, 5) = wsFiches_Action.Cells(i, "P").value 'Statut Rappel
                        Me.ListBox_Rappel.List(Me.ListBox_Rappel.ListCount - 1, 6) = wsFiches_Action.Cells(i, "M").value 'Commentaire
                        
                    End If
                Next i
                
                
                
                 Me.ListBox_Rappel_Histo.Clear
                Me.ListBox_Rappel_Histo.ColumnCount = 6
                Me.ListBox_Rappel_Histo.ColumnWidths = "100;80;80;80;90;80"
                Me.ListBox_Rappel_Histo.BackColor = RGB(200, 200, 255)
                Me.ListBox_Rappel_Histo.Font.Bold = True
                
                ' Ajouter les en-têtes de colonnes
                Me.ListBox_Rappel_Histo.AddItem
                Me.ListBox_Rappel_Histo.List(0, 0) = "Date Rappel"
                Me.ListBox_Rappel_Histo.List(0, 1) = "Creneau"
                Me.ListBox_Rappel_Histo.List(0, 2) = "Motif"
                Me.ListBox_Rappel_Histo.List(0, 3) = "Action"
                Me.ListBox_Rappel_Histo.List(0, 4) = "Statut"
                Me.ListBox_Rappel_Histo.List(0, 5) = "Statut Rappel"
                Me.ListBox_Rappel_Histo.List(0, 6) = "Commentaire"
                
                ' Remplir la ListBox avec les rappels
                For i = 10 To lastRowFichesAction
                    If wsFiches_Action.Cells(i, "E").value = numFiche And wsFiches_Action.Cells(i, "L").value = "Rappel" And wsFiches_Action.Cells(i, "P").value = "Prévu" Then
                        Me.ListBox_Rappel_Histo.AddItem
                        Me.ListBox_Rappel_Histo.List(Me.ListBox_Rappel_Histo.ListCount - 1, 0) = Format(wsFiches_Action.Cells(i, "N").value, "dddd dd/mmmm")
                        Me.ListBox_Rappel_Histo.List(Me.ListBox_Rappel_Histo.ListCount - 1, 1) = Format(wsFiches_Action.Cells(i, "O").value, "hh:mm")
                        Me.ListBox_Rappel_Histo.List(Me.ListBox_Rappel_Histo.ListCount - 1, 2) = wsFiches_Action.Cells(i, "H").value 'Motif
                        Me.ListBox_Rappel_Histo.List(Me.ListBox_Rappel_Histo.ListCount - 1, 3) = wsFiches_Action.Cells(i, "K").value 'Action
                        Me.ListBox_Rappel_Histo.List(Me.ListBox_Rappel_Histo.ListCount - 1, 4) = wsFiches_Action.Cells(i, "L").value 'Statut
                        Me.ListBox_Rappel_Histo.List(Me.ListBox_Rappel_Histo.ListCount - 1, 5) = wsFiches_Action.Cells(i, "P").value 'Statut Rappel
                        Me.ListBox_Rappel_Histo.List(Me.ListBox_Rappel_Histo.ListCount - 1, 6) = wsFiches_Action.Cells(i, "M").value 'Commentaire
                        
                    End If
                Next i
                
                
                
                
                
                ' Vider et configurer la ListBox pour les rappels globaux
                Me.ListBox_Global.Clear
                Me.ListBox_Global.ColumnCount = 5
                Me.ListBox_Global.ColumnWidths = "100;100;100;600;80"
                Me.ListBox_Global.BackColor = RGB(200, 200, 255)
                Me.ListBox_Global.Font.Bold = True
                
                ' Ajouter les en-têtes de colonnes
                Me.ListBox_Global.AddItem
                Me.ListBox_Global.List(0, 0) = "Date Rappel"
                Me.ListBox_Global.List(0, 1) = "Creneau"
                Me.ListBox_Global.List(0, 2) = "Action"
                Me.ListBox_Global.List(0, 3) = "Commentaire"
                Me.ListBox_Global.List(0, 4) = "Motif"
                
                ' Remplir la ListBox avec les rappels globaux
                For i = 10 To lastRowFichesAction
                    If wsFiches_Action.Cells(i, "L").value <> "Clôt" And wsFiches_Action.Cells(i, "L").value = "Rappel" Then
                        Me.ListBox_Global.AddItem
                        Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 0) = Format(wsFiches_Action.Cells(i, "K").value, "dddd dd/mm")
                        Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 1) = Format(wsFiches_Action.Cells(i, "L").value, "hh:mm")
                        Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 2) = wsFiches_Action.Cells(i, "I").value
                        Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 3) = wsFiches_Action.Cells(i, "M").value
                        Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 4) = wsFiches_Action.Cells(i, "H").value
                    End If
                Next i
                
                ' Marquer comme trouvé et quitter la boucle
                found = True
                Me.OptionButton_Existant.value = True
                Me.TextBox_Searsh.Text = ""
                Me.TextBox_Searsh.Visible = False
                Exit For
            End If
        Next cell
        
        ' Si aucune donnée n'est trouvée
        If Not found Then
            Me.OptionButton_Nouveau.value = True
            MsgBox "Aucune donnée trouvée pour le nom : " & nomRecherche, vbInformation
            Me.txtRef.Text = nomRecherche
            wsConsole.Range("E28").value = Me.txtRef.value
        End If
    End If
     Application.EnableEvents = True
End Sub
Private Sub RemoveDuplicatesFromComboBox(combo As MSForms.ComboBox)
    Dim i As Long, j As Long
    For i = combo.ListCount - 1 To 0 Step -1
        For j = 0 To i - 1
            If combo.List(i) = combo.List(j) Then
                combo.RemoveItem i
                Exit For
            End If
        Next j
    Next i
End Sub
Private Sub ComboBox1_Change()

    Dim i As Integer
    Dim txt As String
    txt = Me.ComboBox1.Text
    

    Application.EnableEvents = False
  
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
    
    
    

    Application.EnableEvents = True
End Sub
Sub Global_Stat()

    Dim wsFiches As Worksheet, wsActions_Fiches As Worksheet
    Dim lastRowFiches As Long, lastRowActions As Long
    Dim i As Long, j As Long
    Dim numFiche As String
    Dim maxDate As Date, maxTime As Date, maxRow As Long

    ' Initialisation des feuilles
    Set wsFiches = ThisWorkbook.Sheets("Fiches")
    Set wsActions_Fiches = ThisWorkbook.Sheets("Actions_Fiches")

    ' Trouver la dernière ligne des feuilles
    lastRowFiches = wsFiches.Cells(wsFiches.Rows.Count, "H").End(xlUp).Row
    lastRowActions = wsActions_Fiches.Cells(wsActions_Fiches.Rows.Count, "E").End(xlUp).Row

    ' Effacer et configurer la ListBox principale
    Me.ListBox_Global.Clear
    Me.ListBox_Global.ColumnCount = 8
    Me.ListBox_Global.ColumnWidths = "80;80;80;80;80;80;80;5"
    Me.ListBox_Global.BackColor = RGB(200, 200, 255)
    Me.ListBox_Global.Font.Bold = True

    ' Ajouter les en-têtes à ListBox_Global
    Me.ListBox_Global.AddItem
    Me.ListBox_Global.List(0, 0) = "Date Rappel"
    Me.ListBox_Global.List(0, 1) = "Creneau"
    Me.ListBox_Global.List(0, 2) = "Assu"
    Me.ListBox_Global.List(0, 3) = "Action"
    Me.ListBox_Global.List(0, 4) = "Commentaire"
    Me.ListBox_Global.List(0, 5) = "Motif"
    Me.ListBox_Global.List(0, 6) = "Résumé"
    Me.ListBox_Global.List(0, 7) = "NumFiche"

    ' Parcourir la feuille "Fiches" pour remplir ListBox_Global
    For i = 11 To lastRowFiches
        If wsFiches.Cells(i, "N").value <> "Clôt" Then
            numFiche = wsFiches.Cells(i, "H").value

            ' Ajouter l'enregistrement à ListBox_Global
            Me.ListBox_Global.AddItem
            Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 0) = Format(wsFiches.Cells(i, "E").value, "dddd dd/mm")
            Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 1) = Format(wsFiches.Cells(i, "D").value, "hh:mm")
            Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 2) = wsFiches.Cells(i, "F").value
            Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 3) = wsFiches.Cells(i, "L").value
            Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 4) = wsFiches.Cells(i, "O").value
            Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 5) = wsFiches.Cells(i, "L").value
            Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 6) = wsFiches.Cells(i, "N").value
            Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 7) = numFiche
        End If
    Next i

    ' Effacer et configurer la ListBox des rappels
    Me.ListBox_Global_Rappels.Clear
    Me.ListBox_Global_Rappels.ColumnCount = 8
    Me.ListBox_Global_Rappels.ColumnWidths = "80;80;80;80;80;80;80;5"
    Me.ListBox_Global_Rappels.BackColor = RGB(200, 200, 255)
    Me.ListBox_Global_Rappels.Font.Bold = True

    ' Ajouter les en-têtes à ListBox_Global_Rappels
    Me.ListBox_Global_Rappels.AddItem
    Me.ListBox_Global_Rappels.List(0, 0) = "Date Rappel"
    Me.ListBox_Global_Rappels.List(0, 1) = "Creneau"
    Me.ListBox_Global_Rappels.List(0, 2) = "Assu"
    Me.ListBox_Global_Rappels.List(0, 3) = "Action"
    Me.ListBox_Global_Rappels.List(0, 4) = "Commentaire"
    Me.ListBox_Global_Rappels.List(0, 5) = "Motif"
    Me.ListBox_Global_Rappels.List(0, 6) = "Statut"
    Me.ListBox_Global_Rappels.List(0, 7) = "NumFiche"

    ' Parcourir chaque NumFiche dans la feuille Fiches
    For i = 11 To lastRowFiches
        numFiche = wsFiches.Cells(i, "H").value ' NumFiche dans la colonne H

        ' Initialiser les variables pour stocker la date et l'heure maximales
        maxDate = DateSerial(1900, 1, 1)
        maxTime = TimeSerial(0, 0, 0)
        maxRow = -1

        ' Parcourir la feuille Actions_Fiches pour trouver les rappels correspondants
        For j = 10 To lastRowActions
            If wsActions_Fiches.Cells(j, "E").value = numFiche And _
               wsActions_Fiches.Cells(j, "K").value <> "Clôt" And _
               wsActions_Fiches.Cells(j, "L").value = "Rappel" Then

                ' Vérifier si la date et l'heure sont plus récentes
                If wsActions_Fiches.Cells(j, "N").value > maxDate Or _
                   (wsActions_Fiches.Cells(j, "N").value = maxDate And wsActions_Fiches.Cells(j, "O").value > maxTime) Then
                    maxDate = wsActions_Fiches.Cells(j, "N").value
                    maxTime = wsActions_Fiches.Cells(j, "O").value
                    maxRow = j ' Enregistrer la ligne du rappel le plus récent
                End If
            End If
        Next j

        ' Si un rappel a été trouvé, ajouter les informations à la ListBox
        If maxRow <> -1 Then
            Me.ListBox_Global_Rappels.AddItem
            Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 0) = Format(wsActions_Fiches.Cells(maxRow, "N").value, "dddd dd/mm")
            Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 1) = Format(wsActions_Fiches.Cells(maxRow, "O").value, "hh:mm")
            Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 2) = wsActions_Fiches.Cells(maxRow, "C").value
            Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 3) = wsActions_Fiches.Cells(maxRow, "I").value
            Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 4) = wsActions_Fiches.Cells(maxRow, "M").value
            Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 5) = wsActions_Fiches.Cells(maxRow, "H").value
            Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 6) = wsActions_Fiches.Cells(maxRow, "P").value
            Me.ListBox_Global_Rappels.List(Me.ListBox_Global_Rappels.ListCount - 1, 7) = numFiche
        End If
    Next i

    ' Mettre à jour l'onglet
    UpdateTabCaption

End Sub

Private Sub UpdateTabCaption()
    Dim itemCount As Integer
    ' Accéder à la ListBox sur Page2 et compter les éléments
    itemCount = Me.ListBox_Global_Rappels.ListCount
    
    ' Mettre à jour le nom de l'onglet de Page2 dans le MultiPage
    Me.MultiPage1.Pages(1).Caption = "Rappel (" & itemCount & ")"
End Sub
Private Sub ListBox_Global_Rappels_Change()
    UpdateTabCaption
End Sub

'Private Sub MultiPage1_Change()
    'If Me.MultiPage1.value = 1 Then ' Si l'utilisateur passe à Page2
       ' UpdateTabCaption
   ' End If
'End Sub

Private Sub ComboBox_Change()
    Dim txt As String
    Dim i As Integer
    Dim uniqueItems As Collection
    Dim item As Variant
    Dim filteredList As String
    Dim wsDatabase As Worksheet
    Dim lastRowDatabase As Long
    
    ' Désactiver les événements pour éviter des boucles infinies
    'Application.EnableEvents = False
    
    ' Récupérer le texte saisi dans ComboBox1
    txt = Me.ComboBox1.Text
    
    ' Effacer la liste actuelle de ComboBox1
    Me.ComboBox1.Clear
    
    ' Définir la feuille "database"
    Set wsDatabase = ThisWorkbook.Sheets("Fiches")
    
    ' Trouver la dernière ligne de la colonne D dans "database"
    lastRowDatabase = wsDatabase.Cells(wsDatabase.Rows.Count, "J").End(xlUp).Row
    
    ' Créer une collection pour stocker les éléments uniques
    Set uniqueItems = New Collection
    
    ' Parcourir la colonne D pour trouver les éléments correspondants
    For i = 10 To lastRowDatabase
        If LCase(Left(wsDatabase.Cells(i, "J").value, Len(txt))) = LCase(txt) Then
            On Error Resume Next
            uniqueItems.Add wsDatabase.Cells(i, "J").value, CStr(wsDatabase.Cells(i, "J").value)
            On Error GoTo 0
        End If
    Next i
    
    ' Ajouter les éléments uniques à la ComboBox1
    For Each item In uniqueItems
        Me.ComboBox1.AddItem item
    Next item
    
    ' Réactiver les événements
    Application.EnableEvents = True
End Sub



Private Sub TextBox_Searsh_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Vérifie si l'utilisateur appuie sur la touche Entrée
    If KeyCode = vbKeyReturn Then
        Dim wsDatabase As Worksheet, wsConsole As Worksheet
        Dim lastRowDatabase As Long
        Dim nomRecherche As String
        Dim numFiche As String
        Dim found As Boolean
        Dim i As Long, j As Long ' Utiliser une autre variable pour les boucles internes
        
        ' Définir les feuilles
        Set wsDatabase = ThisWorkbook.Sheets("Fiches")
        Set wsConsole = ThisWorkbook.Sheets("Console")
        
        ' Trouver la dernière ligne de la colonne C dans "database"
        lastRowDatabase = wsDatabase.Cells(wsDatabase.Rows.Count, "J").End(xlUp).Row
        
        ' Récupérer le numéro de fiche saisi dans ComboBox1
        numFiche = ExtractNumFiche(Me.TextBox_Searsh.Text)
        
        ' Vérifier si un numéro de fiche a été trouvé
        If numFiche = "" Then
            MsgBox "Le format du texte saisi est incorrect. Veuillez utiliser le format : 'Fiche N°XXXXXXX - Nom'.", vbExclamation
            Exit Sub
        End If
        
        ' Initialiser la variable found
        found = False
        
        ' Parcourir la colonne C pour trouver le numéro de fiche
        For i = 2 To lastRowDatabase ' Commence à la ligne 2 (en supposant que la ligne 1 est l'en-tête)
            If wsDatabase.Cells(i, "H").value = numFiche Then
                ' Récupérer le nom associé au numéro de fiche
                nomRecherche = wsDatabase.Cells(i, "F").value ' Colonne D (Nom Prénom)
                
                ' Afficher la date et la référence
                Me.LabelDate.Caption = Format(wsDatabase.Cells(i, "E").value, "dddd dd mmmm yyyy") ' Colonne E (Now)
                Me.LabelRef.Caption = wsDatabase.Cells(i, "F").value ' Colonne F (Ref)
                Me.LabelRef1.Caption = nomRecherche ' Afficher le nom dans LabelRef1
                
                ' Effacer la liste existante dans List_Histo
                Me.List_Histo.Clear
                
                ' Configurer la ListBox pour afficher des colonnes
                Me.List_Histo.ColumnCount = 4 ' Nombre de colonnes (Date, Heure, Commentaire, Résumé)
                Me.List_Histo.ColumnWidths = "100;60;600;100" ' Largeur des colonnes
                Me.List_Histo.BackColor = RGB(200, 200, 255) ' Couleur de fond (bleu clair)
                Me.List_Histo.Font.Bold = True ' Texte en gras
                
                ' Ajouter les en-têtes à List_Histo
                Me.List_Histo.AddItem
                Me.List_Histo.List(0, 0) = "Date"
                Me.List_Histo.List(0, 1) = "Heure"
                Me.List_Histo.List(0, 2) = "Commentaire"
                Me.List_Histo.List(0, 3) = "Résumé"
                
                ' Parcourir la colonne C de "database" pour trouver les enregistrements correspondants au Num Fiche
                For j = 2 To lastRowDatabase ' Utiliser une autre variable (j) pour la boucle interne
                    If wsDatabase.Cells(j, "C").value = numFiche Then ' Colonne C (Num Fiche)
                        ' Ajouter l'enregistrement à List_Histo
                        Me.List_Histo.AddItem
                        Me.List_Histo.List(Me.List_Histo.ListCount - 1, 0) = Format(wsDatabase.Cells(j, "E").value, "dddd dd mmmm yyyy") ' Date (colonne E)
                        Me.List_Histo.List(Me.List_Histo.ListCount - 1, 1) = Format(wsDatabase.Cells(j, "D").value, "hh:mm") ' Heure (colonne E)
                        Me.List_Histo.List(Me.List_Histo.ListCount - 1, 2) = wsDatabase.Cells(j, "O").value ' Commentaire (colonne G)
                        Me.List_Histo.List(Me.List_Histo.ListCount - 1, 3) = wsDatabase.Cells(j, "N").value ' Résumé (colonne H)
                    End If
                Next j
                
              
                Me.ListBox_Rappel.Clear
                
               
                Me.ListBox_Rappel.ColumnCount = 5 '
                Me.ListBox_Rappel.ColumnWidths = "100;100;100;600;80" '
                Me.ListBox_Rappel.BackColor = RGB(200, 200, 255) '
                Me.ListBox_Rappel.Font.Bold = True
           
                Me.ListBox_Rappel.AddItem
                Me.ListBox_Rappel.List(0, 0) = "Date Rappel"
                Me.ListBox_Rappel.List(0, 1) = "Creneau"
                Me.ListBox_Rappel.List(0, 2) = "Action"
                Me.ListBox_Rappel.List(0, 3) = "Commentaire"
                Me.ListBox_Rappel.List(0, 4) = "Motif"
                
                
                For j = 2 To lastRowDatabase '
                    If wsDatabase.Cells(j, "C").value = numFiche And wsDatabase.Cells(j, "H").value = "Rappel" Then '
                       
                        Me.ListBox_Rappel.AddItem
                        Me.ListBox_Rappel.List(Me.ListBox_Rappel.ListCount - 1, 0) = Format(wsDatabase.Cells(j, "E").value, "dddd dd/mm/yyyy") '
                        Me.ListBox_Rappel.List(Me.ListBox_Rappel.ListCount - 1, 1) = Format(wsDatabase.Cells(j, "E").value, "hh:mm") '
                        Me.ListBox_Rappel.List(Me.ListBox_Rappel.ListCount - 1, 2) = wsDatabase.Cells(j, "I").value
                        Me.ListBox_Rappel.List(Me.ListBox_Rappel.ListCount - 1, 3) = wsDatabase.Cells(j, "J").value
                        Me.ListBox_Rappel.List(Me.ListBox_Rappel.ListCount - 1, 4) = wsDatabase.Cells(j, "K").value '
                    End If
                Next j
                
               
                Me.ListBox_Global.Clear
                Me.ListBox_Global.ColumnCount = 5 '
                Me.ListBox_Global.ColumnWidths = "100;100;100;600;80" '
                Me.ListBox_Global.BackColor = RGB(200, 200, 255) '
                Me.ListBox_Global.Font.Bold = True
                
                
                Me.ListBox_Global.AddItem
                Me.ListBox_Global.List(0, 0) = "Date Rappel"
                Me.ListBox_Global.List(0, 1) = "Creneau"
                Me.ListBox_Global.List(0, 2) = "Action"
                Me.ListBox_Global.List(0, 3) = "Commentaire"
                Me.ListBox_Global.List(0, 4) = "Motif"
                
           
                For j = 2 To lastRowDatabase ' Utiliser une autre variable (j) pour la boucle interne
                    If wsDatabase.Cells(j, "C").value = numFiche And wsDatabase.Cells(j, "H").value = "Rappel" Then ' Colonne C (Num Fiche) et Colonne H (Résumé)
                        ' Ajouter l'enregistrement à Listbox_Global
                        Me.ListBox_Global.AddItem
                        Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 0) = Format(wsDatabase.Cells(j, "E").value, "dddd dd/mm/yyyy") ' Date Rappel (colonne E)
                        Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 1) = Format(wsDatabase.Cells(j, "E").value, "hh:mm") ' Creneau (colonne E)
                        Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 2) = wsDatabase.Cells(j, "I").value ' Motif (colonne I)
                        Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 3) = wsDatabase.Cells(j, "J").value ' Action (colonne J)
                        Me.ListBox_Global.List(Me.ListBox_Global.ListCount - 1, 4) = wsDatabase.Cells(j, "K").value ' Commentaire (colonne K)
                    End If
                Next j
                
                found = True
                Exit For ' Sortir dès qu'on trouve le numéro de fiche
            End If
        Next i
        
        Me.MultiPage1.value = 0
        OptionButton_Existant.value = True
        
        ' Vérifier si des données ont été trouvées
        If Not found Then
            OptionButton_Nouveau.value = True
            MsgBox "Aucune donnée trouvée pour le numéro de fiche : " & numFiche, vbInformation
            txtRef.Text = numFiche
            wsConsole.Range("E28").value = Me.txtRef.value
        End If
    End If
End Sub

Private Function ExtractNumFiche(comboText As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim numFiche As String
    
    ' Créer une expression régulière pour extraire le numéro de fiche
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "Fiche N°(\d+)"
    regex.IgnoreCase = True
    regex.Global = True
    
    ' Appliquer l'expression régulière au texte
    Set matches = regex.Execute(comboText)
    
    ' Si un numéro de fiche est trouvé, le retourner
    If matches.Count > 0 Then
        numFiche = matches(0).SubMatches(0)
    Else
        numFiche = "" ' Retourner une chaîne vide si aucun numéro n'est trouvé
    End If
    
    ExtractNumFiche = numFiche
End Function

