Attribute VB_Name = "AttachmentHelpers"
Option Compare Database
Option Explicit

Public Function ConvertFileToBase64(sPath As String) As String
    Dim bytes
    With CreateObject("ADODB.Stream")
    .Open
    .Type = 1 'ADODB.adTypeBinary
    .LoadFromFile sPath
    bytes = .Read
    .Close
    End With
    ConvertFileToBase64 = EncodeBase64(bytes)
End Function

Private Function EncodeBase64(bytes) As String

    Dim objXML                      As MSXML2.DOMDocument60
    Dim objNode                     As MSXML2.IXMLDOMElement

    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = bytes
    EncodeBase64 = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing
End Function

