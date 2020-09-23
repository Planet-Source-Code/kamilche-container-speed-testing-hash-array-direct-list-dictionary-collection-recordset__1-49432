VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' IContainer - Interface for all the container classes
'--------------------------------------------------------------------------------
Option Explicit

Private Const module As String = "IContainer"   '$ PROBHIDE DEAD

'$ PROBHIDE NO_ERROR_HANDLER BEGIN

Public Property Get Name() As String
    ' Return the name of the container
End Property

Public Function Add(ByVal Data As Variant, ByVal Key As Long) As Long
    ' Add an item to the container
End Function

Public Property Get Item(ByVal Key As Long) As Variant  '$ PROBHIDE EXCESS_SCOPE
Attribute Item.VB_UserMemId = 0
    ' Get an item from the container
End Property

Public Property Let Item(ByVal Key As Long, ByVal Data As Variant)
    ' Set an item in the container
End Property

Public Function Lookup(ByVal Key As Long, ByRef Data As Variant) As Boolean
    ' Lookup an item from the container
End Function

Public Function Remove(ByVal Key As Long) As Boolean
    ' Remove an item from the container
End Function

Public Sub Clear() '$ PROBHIDE EXCESS_SCOPE
    ' Remove all items from the container
End Sub

Public Function Exists(ByVal Key As Long) As Boolean
    ' Return whether or not an item exists in the container
End Function

Public Function Count() As Long
    ' Return the item count
End Function

Public Function Keys() As Long()
    ' Return an array of all keys in the container
End Function

Public Function GetTimings(ByVal obj As IContainer, ByVal Loops As Long) As clsTimings
    ' Time this class
    On Error GoTo Err_Init
    Dim i As Long
    Dim Timer As clsTimer
    Dim Keys() As Long
    Dim Data As Long
    Dim mTimings As clsTimings
        
    Set mTimings = New clsTimings
    Set Timer = New clsTimer
    
    Screen.MousePointer = vbHourglass
    
    ' Test adds
    Timer.StartTimer
    For i = 1 To Loops
        If obj.Add(i + 100, i) = 0 Then
            MsgBox "Error!"
            Exit For
        End If
    Next i
    Timer.StopTimer Loops
    mTimings.AddsPerSecond = Timer.LoopsPerSecond
    mTimings.BytesPerLoop = Timer.BytesPerLoop
    
    If obj.Name <> "clsList" Then
        ' Test exists
        Timer.StartTimer
        For i = 1 To Loops
            If obj.Exists(i) = False Then
                MsgBox "Error!"
                Exit For
            End If
        Next i
        Timer.StopTimer Loops
        mTimings.ExistsPerSecond = Timer.LoopsPerSecond
        
        ' Test gets
        Timer.StartTimer
        For i = 1 To Loops
            If obj.Item(i) <> i + 100 Then
                MsgBox "Error!"
                Exit For
            End If
        Next i
        Timer.StopTimer Loops
        mTimings.GetsPerSecond = Timer.LoopsPerSecond
        
        ' Test sets
        Timer.StartTimer
        For i = 1 To Loops
            obj.Item(i) = i + 101
        Next i
        Timer.StopTimer Loops
        mTimings.SetsPerSecond = Timer.LoopsPerSecond
        For i = 1 To Loops
            If obj.Item(i) <> i + 101 Then
                MsgBox "Error!"
                Exit For
            End If
        Next i
    
        ' Test lookups
        Timer.StartTimer
        For i = 1 To Loops
            obj.Lookup i, Data
            If Data <> i + 101 Then
                MsgBox "Error!"
                Exit For
            End If
        Next i
        Timer.StopTimer Loops
        mTimings.LookupsPerSecond = Timer.LoopsPerSecond
    End If
    
    ' Test keys
    Timer.StartTimer
    For i = 1 To 1
        Keys = obj.Keys
    Next i
    Timer.StopTimer 1
    mTimings.KeysPerSecond = Timer.LoopsPerSecond
    For i = 0 To Loops - 1
        If Keys(i) <> i + 1 Then
            MsgBox "Error!"
            Exit For
        End If
    Next i
    
    ' Test removes
    Timer.StartTimer
    For i = 1 To Loops
        If obj.Remove(i) = False Then
            MsgBox "Error!"
            Exit For
        End If
    Next i
    Timer.StopTimer Loops
    mTimings.DeletesPerSecond = Timer.LoopsPerSecond
    
    Set GetTimings = mTimings
    Screen.MousePointer = vbDefault
    Exit Function

Err_Init:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Test(ByVal table As IContainer) As Boolean
    ' Test the class to make sure it's working as intended
    On Error GoTo Err_Init
    Dim obj As Collection
    Dim s As String
    Dim n As Long
    Dim Keys() As Long
    Dim i As Long
    
    Test = False
    
    ' Create test object
    Set obj = New Collection
    obj.Add "cindy", "kam"
    s = obj.Item("kam")
    If s <> "cindy" Then Exit Function
    
    ' Add items to table
    If table.Add(obj, 1) = False Then Exit Function
    If table.Add(5, 2) = False Then Exit Function
    If table.Add("hi", 3) = False Then Exit Function
    If table.Add(Null, 4) = False Then Exit Function
    If table.Count <> 4 Then Exit Function
    
    ' Test object lookup
    Set obj = Nothing
    If table.Name <> "clsList" Then
        Set obj = table.Item(1)
        If obj.Item("kam") <> "cindy" Then Exit Function
    End If
    Set obj = Nothing
    If table.Exists(999) = True Then Exit Function
    If table.Exists(1) = False Then Exit Function
    If table.Name <> "clsList" Then
        If table.Lookup(1, obj) = False Then Exit Function
    End If
    
    ' Test valid lookups
    If table.Name <> "clsList" Then
        Set obj = table(1)
        If obj Is Nothing Then Exit Function
        n = table(2)
        If n <> 5 Then Exit Function
        s = table(3)
        If s <> "hi" Then Exit Function
        If Not IsNull(table(4)) Then Exit Function
    End If
    
    ' Test adding duplicates
    If table.Name <> "clsList" And table.Name <> "clsDirect" Then
        If table.Add(Null, 1) <> 0 Then Exit Function
        If table.Add(Null, 2) <> 0 Then Exit Function
        If table.Add(Null, 3) <> 0 Then Exit Function
        If table.Add(Null, 4) <> 0 Then Exit Function
    Else
        If table.Add(Null, 1) = 0 Then Exit Function
    End If
    
    ' Test the key list
    Keys = table.Keys
    For i = 1 To 4
        If Keys(i - 1) <> i Then Exit Function
    Next i
    
    ' Test clearing the table
    table.Clear
    If table.Count <> 0 Then Exit Function
   
    Test = True
    
    Exit Function

Err_Init:
    MsgBox Err.Number & " - " & Err.Description
End Function


'$ END
