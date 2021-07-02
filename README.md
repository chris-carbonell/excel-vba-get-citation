# excel_vba_get_citation
Excel function to get citation from DOI URL

# Quick Start

The following:<br>
<code>=GetCitation("https://doi.org/10.1016/j.ijinfomgt.2019.102055")</code>

returns:<br>
> Ghasemaghaei, M. (2021). Understanding the impact of big data on firm performance: The necessity of conceptually differentiating among big data characteristics. International Journal of Information Management, 57, 102055. doi:10.1016/j.ijinfomgt.2019.102055

# Installation

You need to add a reference to <b>Microsoft XML v6.0</b> as follows:
1. In the VBA Editor, navigate to Tools > References.
2. In the References window, add a check to the checkbox for Microsoft XML v6.0.

# VBA

<code>
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
</code>

# Workbook

Personally, I used this function to really speed up my documentation of references during research. See <code>example.xlsm</code>.

You'll see that the workbook also provides the narrative and parenthetical citations (which properly account for any number of authors) as well!

# Knowledge
* Excel VBA
* HTTP requests

# Warnings

* Since the VBA is making a request every time the cell is calculated, I'd recommend that you enter the formula, get the results, and then copy and paste as values to save the results.
* Depending on your professor, the resulting references are <b>NOT</b> perfect. Review them to ensure compliance with your institution's requirements.

# Resources
* Loose Tutorial
  * https://citation.crosscite.org/docs.html
* Very Helpful Code
  * https://stackoverflow.com/questions/6984528/sending-form-data-through-xmlhttp-in-vba
  * https://stackoverflow.com/questions/22938194/xmlhttp-request-is-raising-an-access-denied-error
* Testing
  * https://reqbin.com/
