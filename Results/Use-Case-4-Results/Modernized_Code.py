def sort_multiple_columns():
    ws = Excel.ActiveSheet
    sort_fields = ws.Sort.SortFields
    
    sort_fields.Add(Key=ws.Range("A1"), SortAscending=True)
    sort_fields.Add(Key=ws.Range("B1"), SortAscending=True)
    
    sort_fields.SetRange(ws.Range("A1:C13"))
    ws.Sort.Header = xlYes
    ws.Sort.Apply()
