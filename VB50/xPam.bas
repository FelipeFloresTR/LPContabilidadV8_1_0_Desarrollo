Attribute VB_Name = "xPam"
' Para Microsoft XML 3.0

'   Dim xIAsist As IXMLDOMNode, xAsist As IXMLDOMNode
'   Dim xADoc As MSXML2.DOMDocument
'
'   Set xADoc = New MSXML2.DOMDocument
'
'   Set xIAsist = xADoc.createElement("InfoAsist")
'   Call xADoc.appendChild(xIAsist)
'
'   Call xAddTag(xADoc, xIAsist, "Version", xAVersion) ' version del xml, por si despues cambia
'
'   Set xAsist = xAddTag(xADoc, xIAsist, "Asistencia")
'
'   Call xAddTag(xADoc, xAsist, "PcName", gEnv.PC_Name)
'   Call xAddTag(xADoc, xAsist, "UsrName", gEnv.UserName)
'

Public Function xAddTag(xDoc As MSXML2.DOMDocument, xPadre As IXMLDOMNode, ByVal TagName As String, Optional ByVal TagValue As String = "") As IXMLDOMNode
   Dim xNodo As IXMLDOMNode

   Set xNodo = xDoc.createElement(Trim(TagName))
   
   If Trim(TagValue) <> "" Then
      xNodo.nodeTypedValue = Trim(TagValue)
   End If
  
   Call xPadre.appendChild(xNodo)

   Set xAddTag = xNodo
End Function

Public Function xAddAttrib(xDoc As MSXML2.DOMDocument, xNodo As IXMLDOMNode, ByVal AttrName As String, Optional ByVal AttrValue As String = "") As IXMLDOMAttribute
   Dim xAttr As IXMLDOMAttribute

   Set xAttr = xDoc.createAttribute(AttrName)
   'xAttr.Value = AttrValue
   xAttr.Text = AttrValue
   Set xAddAttrib = xAttr
   
End Function

Public Function xAddTagWithAttr(xDoc As MSXML2.DOMDocument, xPadre As IXMLDOMNode, ByVal TagName As String, Optional ByVal TagValue As String = "", Optional ByVal AttrName As String, Optional ByVal AttrValue As String = "") As IXMLDOMNode
   Dim xNodo As IXMLDOMNode
   Dim xAttr As IXMLDOMAttribute
   Dim objElement As MSXML2.IXMLDOMElement

   Set xNodo = xAddTag(xDoc, xPadre, TagName, TagValue)
   Set xAttr = xAddAttrib(xDoc, xNodo, AttrName, AttrValue)
   Set objElement = xNodo
   objElement.setAttributeNode xAttr
   Set xAttr = Nothing

   Set xAddTagWithAttr = xNodo
End Function

