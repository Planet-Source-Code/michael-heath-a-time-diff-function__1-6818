Attribute VB_Name = "Module1"
Public StartTime As String      ' Holds the begining time in short format

Public Function GetMins(tm1 As String, tm2 As Date) As Long
    Dim m1 As Long, m2 As Long
    Dim strm1 As String
    
    strm1 = Right(tm1, InStr(tm1, ":") - 1)
    m1 = strm1
   
    m2 = Minute(tm2)
    
    
    GetMins = m2 - m1

End Function
Public Function GetHours(tm1 As String, tm2 As Date) As Long
    
    Dim h1 As Long, h2 As Long


    Dim strh1 As String
    strh1 = Left(tm1, InStr(tm1, ":") - 1)
    h1 = strh1
    
    h2 = Hour(tm2)
        
    GetHours = h2 - h1

End Function

