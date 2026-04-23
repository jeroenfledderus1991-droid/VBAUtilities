Attribute VB_Name = "VBAFormatterContinuationTest"
Option Explicit

' Importeer deze module in de VBA Editor en run daarna de formatter handmatig.
' Controleer vooral dat regels met line continuation (_) als 1 blok bij elkaar blijven.

Public Sub Test_MultilineDimBlock()
    Dim currentValue As String, _
        nextValue As String, _
        finalValue As String

    currentValue = "A"
    nextValue = "B"
    finalValue = "C"

    Debug.Print currentValue & nextValue & finalValue
End Sub

Public Sub Test_MultilineAssignmentBlock()
    Dim messageText As String

    messageText = "De formatter moet dit statement" & _
                  " volledig bij elkaar laten staan" & _
                  " zonder extra lege regel ertussen."

    Debug.Print messageText
End Sub

Public Sub Test_MultilineCallBlock()
    Dim outputPath As String

    outputPath = BuildTestPath( _
        "C:\Temp", _
        "Formatter", _
        "ContinuationTest.txt")

    Debug.Print outputPath
End Sub

Private Function BuildTestPath(ByVal rootFolder As String, _
                               ByVal childFolder As String, _
                               ByVal fileName As String) As String
    BuildTestPath = rootFolder & "\\" & childFolder & "\\" & fileName
End Function