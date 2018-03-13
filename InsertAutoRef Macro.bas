Attribute VB_Name = "NewMacros"
Sub InsertAutoXRef()

Dim sel As Selection
Dim doc As Document
Dim vHeadings As Variant
Dim v As Variant
Dim i As Integer

Dim NewSection As String
Dim CurrentSection As String
Dim NewSubSection As String
Dim CurrentSubSection As String
Dim NewSubSubSection As String
Dim CurrentSubSubSection As String

Set sel = Selection
Set doc = Selection.Document

' Exit if selection includes multiple paragraphs
If sel.Range.Paragraphs.Count <> 1 Then Exit Sub

' Collapse selection if there are spaces or paragraph
' marks on either end
sel.MoveStartWhile cset:=(Chr$(32) & Chr$(13)), Count:=sel.Characters.Count
sel.MoveEndWhile cset:=(Chr$(32) & Chr$(13)), Count:=-sel.Characters.Count

vHeadings = doc.GetCrossReferenceItems(wdRefTypeNumberedItem)

CurrentSection = ""
CurrentSubSection = ""
CurrentSubSubSection = ""
i = 1

For Each v In vHeadings
    v = Trim(v)
    NewSection = FirstMatch("^\d+.\d+", (v))
    NewSubSection = FirstMatch("^\([a-z]\)", (v))
    NewSubSubSection = FirstMatch("^\([iv]+\)", (v))
     
    If NewSection <> "" Then
        CurrentSection = NewSection
        CurrentSubSection = ""
        CurrentSubSubSection = ""
        NewSubSection = ""
        NewSubSubSection = ""
    End If
    
    If NewSubSection <> "" Then
        Select Case NewSubSection
            Case "(i)"
                If CurrentSubSection = "(h)" Then
                    ' This instance of (i) is preceded by (h) and is NOT a roman numeral. Process as a subsection.
                    CurrentSubSection = NewSubSection
                    NewSubSubSection = ""
                    CurrentSubSubSection = ""
                Else
                ' Not preceded by (h), so this instance of (i) will be captured and dealt with below as a subsubsection--ignore
                End If
            Case "(v)"
                If CurrentSubSection = "(u)" Then
                    ' This instance of (v) is preceded by (u) and is NOT a roman numeral. Process as a subsection.
                    CurrentSubSection = NewSubSection
                    NewSubSubSection = ""
                    CurrentSubSubSection = ""
                Else
                    ' Not preceded by (u), so this instance of (v) will be captured and dealt with below as a subsubsection--ignore
                End If
            Case Else
                CurrentSubSection = NewSubSection
                NewSubSubSection = ""
                CurrentSubSubSection = ""
        End Select
    End If
    
    If NewSubSubSection <> "" Then
        CurrentSubSubSection = NewSubSubSection
    End If
     
    CurrentHeading = CurrentSection & CurrentSubSection & CurrentSubSubSection
    
    'MsgBox CurrentHeading
     
    If Trim(sel.Range.Text) = Trim(CurrentHeading) Then
        sel.InsertCrossReference _
           referencetype:=wdRefTypeNumberedItem, _
           referencekind:=wdNumberFullContext, _
           referenceitem:=i
       Exit Sub
    End If
i = i + 1
Next v

MsgBox "Couldn't match: " & sel.Range.Text
End Sub

Function FirstMatch(myPattern As String, myString As String)
   'Create objects.
   Dim objRegExp As RegExp
   Dim objMatch As Match
   Dim colMatches As MatchCollection
   Dim RetStr As String
   
   ' Create a regular expression object.
   Set objRegExp = New RegExp

   'Set the pattern by using the Pattern property.
   objRegExp.Pattern = myPattern

   ' Set Case Insensitivity.
   objRegExp.IgnoreCase = True

   'Set global applicability.
   objRegExp.Global = True

   'Test whether the String can be compared.
   If (objRegExp.Test(myString) = True) Then

   'Get the matches.
    Set colMatches = objRegExp.Execute(myString)   ' Execute search.

    RetStr = colMatches(0).Value
    
   Else
    RetStr = ""
   End If
   FirstMatch = RetStr
End Function
