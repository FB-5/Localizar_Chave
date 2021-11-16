Option Explicit

Dim objDoc
Dim objNo

Set objDoc = CreateObject("MSXML2.DOMDocument")

objDoc.load(".\Teste.xml")


Set objNo = objDoc.selectSingleNode("//Teste/Nome")

MsgBox objNo.text