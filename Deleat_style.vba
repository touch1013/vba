Sub delete_style()

    On Error Resume Next

'書式（スタイル）定義を全削除

    Dim M()

    J = ActiveWorkbook.Styles.Count
    ReDim M(J)
    For i = 1 To J
        M(i) = ActiveWorkbook.Styles(i).Name
    Next
    For i = 1 To J
        If InStr("Hyperlink,Normal,Followed Hyperlink", M(i)) = 0 Then
            ActiveWorkbook.Styles(M(i)).Delete
        End If
    Next

End Sub
