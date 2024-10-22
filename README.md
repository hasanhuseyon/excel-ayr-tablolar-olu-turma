Sub CreateMultipleTables()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblRange As Range
    Dim i As Integer
    
    ' Sayfa1'i ayarla
    Set ws = ThisWorkbook.Sheets("Sayfa1")
    
    ' Belirtilen aralıklar için tablo oluştur
    Dim ranges As Variant
    Dim tableNames As Variant
    
    ' Aralıklar ve tablo isimlerini belirle
    ranges = Array("A2:L61", "A62:L77", "A78:L102", "A103:L117", "A118:L141", _
                   "A142:L169", "A170:L187", "A188:L191", "A192:L206", _
                   "A207:L209", "A210:L213", "A214:L217", "A218:L226", "A227:L229", "A230:L230")
    tableNames = Array("Table1", "Table2", "Table3", "Table4", "Table5", _
                       "Table6", "Table7", "Table8", "Table9", _
                       "Table10", "Table11", "Table12", "Table13", "Table14", "Table15")
    
    ' Her aralık için tablo oluşturma döngüsü
    For i = LBound(ranges) To UBound(ranges)
        Set tblRange = ws.Range(ranges(i))
        Set tbl = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        tbl.Name = tableNames(i)
    Next i
    
End Sub
