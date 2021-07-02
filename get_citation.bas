Attribute VB_Name = "get_citation"
' # Resources

' ## Loose Tutorial
' * https://citation.crosscite.org/docs.html

' ## Very Helpful Code
' * https://stackoverflow.com/questions/6984528/sending-form-data-through-xmlhttp-in-vba
' * https://stackoverflow.com/questions/22938194/xmlhttp-request-is-raising-an-access-denied-error

' ## Testing
' * https://reqbin.com/

Function GetCitation(str_doi_url As String) As String
' get citation from DOI url

    Dim xml As MSXML2.XMLHTTP60
    Dim result As String
    
    Set xml = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' get
    With xml
      .Open "GET", str_doi_url, False
      .setRequestHeader "Accept", "text/x-bibliography; style=apa; locale=en-US"
      .send
    End With
    
    GetCitation = xml.responseText

End Function
