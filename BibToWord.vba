' This VBA macro reads BibTeX formatted references from a text field, processes the input,
' and converts it into XML format to add it as a source in Word's bibliography.
' The macro handles different BibTeX entry types (e.g., @book, @article) and extracts key fields 
' like title, year, volume, pages, and author. It processes multiple authors and splits them 
' into first, last, and middle names. Finally, it adds the XML source to the document's bibliography 
' and inserts a citation at the current selection.
' 
' Current code limitations:
'    - This code cannot interpret & (ampersand) symbol and required 'and' in BibTex source i.e. Wiley & Sons
'    - This code cannot interpret other language characters
' Please feel free to update and improve the code. :)
' 
' Author: Aleem Akhtar
' Version: 1.0.0


Private Sub CommandButton1_Click()
    ' Declare variables to store the XML input, tag, and lines
    Dim strXml As String
    Dim strTag As String
    Dim strLine() As String
    
    ' Store the XML content from a text field
    strXml = xmlSource.Text
    
    ' Split the XML content by line breaks into an array
    strLine = Split(strXml, vbCrLf)

    ' Extract the BibTeX type (e.g., @book, @article) and tag from the first line
    Dim category() As String
    category = Split(strLine(0), "{")
    Dim bibType As String
    bibType = category(0)
    Dim bibTag As String
    bibTag = category(1)
    Dim key As String
    Dim value As String
    
    ' Initialize the XML string for the bibliography source
    Dim xml As String
    xml = "<b:Source>"
    
    ' Remove the closing curly brace from the bibTag and add it to XML
    bibTag = Left(bibTag, Len(bibTag) - 1)
    xml = xml & "<b:Tag>" & Trim(bibTag) & "</b:Tag>"
    
    ' Map BibTeX types to corresponding source types in XML
    If (bibType = "@inproceedings") Then
        xml = xml & "<b:SourceType>ConferenceProceedings</b:SourceType>"
    ElseIf (bibType = "@Book" Or bibType = "@book") Then
        xml = xml & "<b:SourceType>BookSection</b:SourceType>"
    ElseIf (bibType = "@article") Then
        xml = xml & "<b:SourceType>JournalArticle</b:SourceType>"
    ElseIf (bibType = "@misc") Then
        xml = xml & "<b:SourceType>Misc</b:SourceType>"
    ElseIf (bibType = "@TechReport") Then
        xml = xml & "<b:SourceType>Report</b:SourceType>"
    Else
        xml = xml & "<b:SourceType>Misc</b:SourceType>"
    End If

    ' Loop through each line of the input, process it, and add to the XML string
    For Each line In strLine
        
        ' Trim leading/trailing spaces from each line
        line = Trim(line)
        
        ' Remove trailing commas except for the last line
        If (Right(line, 1) = ",") Then
            line = Left(line, Len(line) - 1)
        End If
        
        ' Split each line into a key-value pair using '='
        category = Split(line, "=")
        
        ' Check if we have a valid key-value pair
        If (UBound(category) - LBound(category) + 1 = 2) Then
            key = LCase(Trim(category(0))) ' Convert the key to lowercase
            value = Trim(category(1)) ' Extract the value and trim whitespace
            value = Mid(value, 2, Len(category(1)) - 2) ' Remove quotes from the value
            
            ' Map BibTeX keys to corresponding XML fields
            If (key = "title") Then
                xml = xml & ("<b:Title>" & value & "</b:Title>")
            ElseIf (key = "year") Then
                xml = xml & ("<b:Year>" & value & "</b:Year>")
            ElseIf (key = "volume") Then
                xml = xml & ("<b:Volume>" & value & "</b:Volume>")
            ElseIf (key = "booktitle" And bibType = "@inproceedings") Then
                xml = xml & ("<b:ConferenceName>" & value & "</b:ConferenceName>")
            ElseIf (key = "booktitle") Then
                xml = xml & ("<b:BookTitle>" & value & "</b:BookTitle>")
            ElseIf (key = "pages") Then
                xml = xml & ("<b:Pages>" & value & "</b:Pages>")
            ElseIf (key = "number") Then
                xml = xml & ("<b:Number>" & value & "</b:Number>")
            ElseIf (key = "journal") Then
                xml = xml & ("<b:JournalName>" & value & "</b:JournalName>")
            ElseIf (key = "publisher") Then
                xml = xml & ("<b:Publisher>" & value & "</b:Publisher>")
            ElseIf (key = "organization") Then
                xml = xml & ("<b:Publisher>" & value & "</b:Publisher>")
            ElseIf (key = "institution") Then
                xml = xml & ("<b:Institution>" & value & "</b:Institution>")
            ElseIf (key = "author") Then
                ' For author, split multiple authors and create name list
                xml = xml & ("<b:Author><b:Author><b:NameList>")
                Dim nameList As String
                nameList = ""
                Dim fullname() As String
                Dim authors() As String
                authors = Split(value, " and ") ' Split multiple authors
                
                ' Process each author's name (last, first, middle)
                For Each Author In authors
                    Author = Trim(Author)
                    fullname = Split(Author, ", ")
                    nameList = nameList & "<b:Person>"
                    nameList = nameList & "<b:Last>" & fullname(0) & "</b:Last>"
                    
                    ' If the author has both first and last names, split and add them
                    If (UBound(fullname) - LBound(fullname) + 1 = 2) Then
                        fullname = Split(fullname(1), " ")
                        If (UBound(fullname) - LBound(fullname) + 1 = 1) Then
                            nameList = nameList & "<b:First>" & fullname(0) & "</b:First>"
                        Else
                            nameList = nameList & "<b:First>" & fullname(0) & "</b:First>"
                            nameList = nameList & "<b:Middle>" & fullname(1) & "</b:Middle>"
                        End If
                    End If
                    
                    nameList = nameList & "</b:Person>" & vbCrLf ' End the person block
                Next
                
                xml = xml & (nameList) ' Add the name list to XML
                xml = xml & ("</b:NameList></b:Author></b:Author>") ' Close author block
            End If
        End If
    Next

    ' Close the source XML block
    xml = xml & ("</b:Source>")
    
    ' Output the generated XML for debugging purposes
    Debug.Print vbCrLf
    Debug.Print xml
    
    ' Add the new bibliography source to the active document's bibliography
    ActiveDocument.Bibliography.Sources.Add xml
    
    ' Insert a citation field in the document using the bibTag
    Selection.Fields.Add Selection.Range, _
        wdFieldCitation, bibTag
    
    ' Unload the form to close it
    Unload Me
End Sub
