Private Function strip(ByVal str As String) As String
  Dim last
  
  For i = 1 To Len(str) Step 1
    If Asc(Mid(str, i, 1)) < 33 Then
      last = i
    End If
  Next i
  
  If last > 0 Then
    strip = Mid(str, 1, last - 1)
  Else
    strip = str
  End If
End Function

Public Function getSMILES(ByVal name As String) As String
  Dim XMLhttp: Set XMLhttp = CreateObject("MSXML2.ServerXMLHTTP")
  XMLhttp.setTimeouts 2000, 2000, 2000, 2000
  XMLhttp.Open "GET", "http://cactus.nci.nih.gov/chemical/structure/" + name + "/smiles", False
  XMLhttp.send

  If XMLhttp.Status = 200 Then
    getSMILES = strip(XMLhttp.responsetext)
  Else
    getSMILES = ""
  End If
End Function

Public Function getInChIKey(ByVal name As String) As String
  Dim XMLhttp: Set XMLhttp = CreateObject("MSXML2.ServerXMLHTTP")
  XMLhttp.setTimeouts 1000, 1000, 1000, 1000
  XMLhttp.Open "GET", "http://cactus.nci.nih.gov/chemical/structure/" + name + "/stdinchikey", False
  XMLhttp.send

  If XMLhttp.Status = 200 Then
    getInChIKey = Mid(strip(XMLhttp.responsetext), 10)
  Else
    getInChIKey = ""
  End If
End Function

Sub TestSMILES()
  MsgBox (getSMILES("benzene"))
End Sub

Sub TestInChIKey()
  MsgBox (getInChIKey("benzene"))
End Sub
