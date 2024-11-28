VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btn_clear_Click() 'clear

Me.txt_srcode.Value = ""
Me.txt_name.Value = ""
Me.comb_college.Value = ""
Me.comb_program.Value = ""


End Sub

Private Sub btn_pic_Click() 'upload picture

Dim picturePath As String
Dim ws As Worksheet
Dim pic As Shape
Dim targetRange As Range

picturePath = Application.GetOpenFilename("Pictures(*.jpg;*.jpeg;*.gif),*.jpg;*.jpeg;*.gif")

Dim shp As Shape
    For Each shp In ActiveSheet.Shapes 'For loop to search and delete existing pictures to avoid multiple pictures in the pic overview
        If Not Intersect(shp.TopLeftCell, Range("pic_overview")) Is Nothing Then shp.Delete
    Next shp

If picturePath <> "False" Then
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set targetRange = ws.Range("pic_overview")
    
        'to resize the image to fit the range which I defined "pic_overview"
        Set pic = ws.Shapes.AddPicture( _
            Filename:=picturePath, _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoCTrue, _
            Left:=targetRange.Left, _
            Top:=targetRange.Top, _
            Width:=targetRange.Width, _
            Height:=targetRange.Height)
    
        pic.LockAspectRatio = msoFalse
        
        pic.Top = targetRange.Top
        pic.Left = targetRange.Left
        pic.Width = targetRange.Width
        pic.Height = targetRange.Height

    End If
End Sub

Private Sub btn_print_Click() 'printable area into pdf

Dim ws As Worksheet
Dim printRange As Range
Set ws = ThisWorkbook.Sheets("Sheet1")
Set printRange = ws.Range("to_print")
printRange.ExportAsFixedFormat Type:=xlTypePDF, Filename:="C:\Users\lanzd\Downloads\PrintedRange.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True

End Sub

Private Sub btn_submit_Click() 'submit information and picture to the ID

Dim ws As Worksheet
Dim sourceRange As Range
Dim targetRange As Range
Dim pic As Shape
Dim newPic As Shape

Range("D15").Value = txt_srcode.Text
Range("D13").Value = txt_name.Text
Range("C16").Value = comb_program.Text
Range("C18").Value = comb_college.Text

Set ws = ThisWorkbook.Sheets("Sheet1")
Set sourceRange = ws.Range("pic_overview")
Set targetRange = ws.Range("D6:E11")
For Each pic In ws.Shapes   'For loop to search for a picture in Range("pic_overview") to copy to Range("D6:E11") or the ID picture.
    If Not Intersect(pic.TopLeftCell, sourceRange) Is Nothing Then
        Set newPic = pic.Duplicate
            With newPic
                .Top = targetRange.Top
                .Left = targetRange.Left
                .Width = targetRange.Width
                .Height = targetRange.Height
            End With
            Exit For
        End If
        Next pic
End Sub

Private Sub comb_college_Change() 'ComboBox college

'defining the variables for the colleges
    Dim CICS_Programs As Variant
    Dim COE_Programs As Variant
    Dim COET_Programs As Variant
    Dim CAFAD_Programs As Variant
    
'storing the programs of each colleges respectively. NO_Programs = [empty array]
    CICS_Programs = Range("H46:H47")
    COE_Programs = Range("H48:H68")
    CAFAD_Programs = Range("H69:H71")
    COET_Programs = Range("H72:H80")
    NO_Programs = Range("P46:P47")
    


    
'Display only the programs of the selected College under the Program ComboBox
    If comb_college.Value = "" Then
        comb_program.Clear      'clear the program combobox after any college is selected to fix the bug where the selected program stays even after selecting a different college
        comb_program.List = NO_Programs
    ElseIf comb_college.Value = "College of Informatics and Computing Sciences" Then
        comb_program.Clear
        comb_program.List = CICS_Programs
    ElseIf comb_college.Value = "College of Engineering" Then
        comb_program.Clear
        comb_program.List = COE_Programs
    ElseIf comb_college.Value = "College of Engineering Technology" Then
        comb_program.Clear
        comb_program.List = COET_Programs
    ElseIf comb_college.Value = "College of Architecture, Fine Arts and Design" Then
        comb_program.Clear
        comb_program.List = CAFAD_Programs
    End If
End Sub

Private Sub comb_program_Change()

End Sub

