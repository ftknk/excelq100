Attribute VB_Name = "Module001"
Option Explicit

Sub Q001()
    Sheet001.Range("A1:C5").Copy Sheet002.Range("A1")
End Sub
Sub Q002()
    Sheet001.Range("A1:C5").Copy
    Sheet002.Range("A1").PasteSpecial xlPasteValues
    Sheet002.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub
Sub Q003()
    '表以外の情報が入っていたらアウトですけれども…
    Range(Range("b2"), Cells.SpecialCells(xlCellTypeLastCell)).ClearContents
End Sub
Sub Q004()
    '最終セルまで＆定数セルを選択（表以外の情報が入っていたらアウト）
    Range(Range("b2"), Cells.SpecialCells(xlCellTypeLastCell)) _
        .SpecialCells(xlCellTypeConstants).ClearContents
End Sub
Sub Q005()
    '表の範囲と行を格納する場所を予約
    Dim rng As Range, r As Range
    '最終セルまでを予約した場所に格納（D列に見出しが入っている前提）
    Set rng = Range(Range("B3"), Cells.SpecialCells(xlCellTypeLastCell))
    
    '表を1行ずつ取り出して処理
    For Each r In rng.Rows '①
        
        'B列またはC列が空欄かどうか確認
        If IsEmpty(r.Cells(1)) Or IsEmpty(r.Cells(2)) Then '③
            'どちらかが空欄だったら何もしない ④
        
        ElseIf IsNumeric(r.Cells(1)) And IsNumeric(r.Cells(2)) Then
            'ついでだから数値チェックもしておく
            '両方とも数値が入っていたら数式を入れ、書式を設定する
            r.Cells(3).Formula = "=B:B*C:C" '数式「B列×C列」
            r.Cells(3).NumberFormatLocal = "\#,##0" '書式「通貨\のカンマ編集」

        Else
            'ここに来たら、どちらかが数値では無かったということです
        
        End If
    
    Next r '②
    
    '== 余談 ==
    '① For Each r ... rng.Rows → rng から1行取り出して r に格納
    '②「Next」で終わらせずに「Next r」にすると安全。
    '③ r.Cells(n) は、左上から右下へと数えて何個目のセルかを示す。
    '④ if文 の「何もしない」は製品でやると怒られるけれども、
    '　 設計書の無い社内向け便利ツールは読みやすいように
    '　 if と else 両方書いてコメント入れるようにしている。
End Sub
Sub Q006()
    '表の範囲と行を格納する場所を予約
    Dim rng As Range, r As Range
    '最終セルまでを予約した場所に格納（D列に見出しが入っている前提）
    Set rng = Range(Range("A2"), Cells.SpecialCells(xlCellTypeLastCell))
    
    '表を1行ずつ取り出して処理
    For Each r In rng.Rows '①
        
        'B列またはC列が空欄かどうか確認
        If r.Cells(1).Text Like "*-*" Then
            '商品コードに"-"が含まれていたら何もしない ⑤
        
        ElseIf IsEmpty(r.Cells(2)) Or IsEmpty(r.Cells(3)) Then '③
            'どちらかが空欄だったら何もしない ④
        
        ElseIf IsNumeric(r.Cells(2)) And IsNumeric(r.Cells(3)) Then
            'ついでだから数値チェックもしておく
            '両方とも数値が入っていたら数式を入れ、書式を設定する
            r.Cells(4).Formula = "=B:B*C:C" '数式「B列×C列」
            r.Cells(4).NumberFormatLocal = "\#,##0" '書式「通貨\のカンマ編集」

        Else
            'ここに来たら、どちらかが数値では無かったということです
        End If
    Next r '②
    
    '== 余談 ==
    '①～④ Q005と同じ
    '⑤ Like 演算子は正規表現とちょっと違うけど便利
    '   判定の順としては④と⑤どちらが先の方が良いのか。
    '   今回は明らかな仕様→暗黙の仕様という事で、⑤→④の順にしました。
End Sub
Sub Q007()
    '表の範囲と行を格納する場所を予約
    Dim rng As Range, r As Range
    '最終セルまでを予約した場所に格納（B列に見出しが入っている前提）
    Set rng = Range(Range("A2"), Cells.SpecialCells(xlCellTypeLastCell))
    
    '表を1行ずつ取り出して処理
    For Each r In rng.Rows
        
        '日付を入れる場所を確保 & 日付だった場合の目印作成
        Dim dt As Date, isTarget As Boolean: isTarget = False
        
        'A列が日付か判定
        If IsEmpty(r.Cells(1)) Then
            '空欄だったら何もしない
        
        ElseIf IsDate(r.Cells(1)) Then
            '日付でした
            dt = r.Cells(1)
            isTarget = True
        
        Else
            '日付が文字列として入力されている可能性を考慮
            On Error Resume Next
                dt = CDate(r.Cells(1))
                isTarget = (Err.Number = 0)
            On Error GoTo 0
        
        End If
        
        '日付だったら末日を入れる
        If isTarget Then
            r.Cells(2).Value = Format(DateSerial(Year(dt), Month(dt) + 1, 0), "'mmdd")
        End If
        
    Next r
    
End Sub
Sub Q008()
    '表の範囲と行を格納する場所を予約
    Dim rng As Range, r As Range
    '最終セルまでを予約した場所に格納（G列に見出しが入っている前提）
    Set rng = Range(Range("B2"), Cells.SpecialCells(xlCellTypeLastCell))
    
    '表を1行ずつ取り出して処理
    For Each r In rng.Rows
        
        '合計点の入れ物作成
        Dim 合計点 As Integer: 合計点 = 0
        '50点未満の目印作成
        Dim 赤点 As Boolean: 赤点 = False
        
        '行の1列目から5列目を合計していく
        Dim i As Long
        For i = 1 To 5
            
            '合計点算出
            合計点 = 合計点 + r.Cells(i).Value
            
            '50点未満判定
            赤点 = 赤点 Or r.Cells(i).Value < 50
        
        Next i
        
        '合格だったら「合格」を設定
        If 合計点 >= 350 And Not 赤点 Then
            r.Cells(6) = "合格"
        End If
        
    Next r
    
End Sub
Sub Q009()
    
    '合格者シート、成績表シートの場所確保
    Dim wsh合格者 As Worksheet, wsh成績表 As Worksheet
    
    '成績表という名前のシートをアクティブにする
    'なかったらExcelから何らかのエラーが出る
    Set wsh成績表 = ActiveWorkbook.Worksheets("成績表")
    
    'エラー無視領域展開
    On Error Resume Next
        
        '「合格者」という名前のシートを呼んでみる
        '無かったらエラーで何も取れないけどエラーは無視される
        Set wsh合格者 = ActiveWorkbook.Sheets("合格者")
    
    'エラー無視領域閉鎖
    On Error GoTo 0
    
    '合格者シートが無事に呼ばれた場合
    If Not wsh合格者 Is Nothing Then
        'ブツを見せてあげる親切心
        wsh合格者.Activate
        '消していいか確認→ダメなら終了
        If vbCancel = MsgBox("既存の「合格者」シートを削除します。", vbOKCancel) Then Exit Sub
        'ダイアログ無視で今のシートは消してしまう
        Application.DisplayAlerts = False
        wsh合格者.Delete
        Application.DisplayAlerts = True
    End If
    
    'まっさらな合格者シートを作る
    Set wsh合格者 = ActiveWorkbook.Worksheets.Add
    wsh合格者.Name = "合格者"
    
    '合格者リスト（タブ区切り）
    Dim 合格者達 As String
    
    '成績表シート内の表と行を格納する場所を予約
    Dim rng As Range, r As Range
    'ついでに合格者数も覚えておかねば
    Dim i As Long: i = 0
    '成績表シート内の最終セルまでを予約した場所に格納
    Set rng = Range(wsh成績表.Range("A2"), wsh成績表.Cells.SpecialCells(xlCellTypeLastCell))
    
    '成績表シート内の表から1行ずつ取り出して処理
    For Each r In rng.Rows
        '合格だったら合格者リストに追加
        If r.Cells(7).Text = "合格" Then
            合格者達 = 合格者達 & r.Cells(1).Text & vbTab
            i = i + 1
        End If
    Next r
    
    '合格者なしの場合、処理終了
    If i = 0 Then Exit Sub
    
    '合格者リストを合格者シートに追加
    With wsh合格者
        '配列的に横並びなので、コピペで縦並びに変更
        .Range("A1").Resize(1, i) = Split(合格者達, vbTab)
        .Range("A1").Resize(1, i).Copy
        .Range("A2").PasteSpecial Transpose:=True
        .Range("A1").EntireRow.Delete
    End With
        
    '見た目を整える
    wsh合格者.Activate
    wsh合格者.Range("A1").Activate
    
    '== 余談 ==
    '「成績表」が無いとExcelのエラーが出ます
    'Excelのエラーがそのまま使えるものは使っています（シート名が表示されるやつとか）
    'Application.DisplayAlerts をオフにするのはちょっと怖い。途中で落ちたら面倒。
    '日本語変数名、違和感あるけど謎英語よりマシだなって…
    '日本語変数名でも頭に「wsh」とか付けておくと補完(Shift+Tab)で入力できるので楽
    
End Sub
Sub Q010()
    
    '受注シートの場所確保
    Dim wsh As Worksheet
    
    '受注という名前のシートを呼ぶ
    Set wsh = ActiveWorkbook.Worksheets("受注")
    
    '受注シート内の表と行を格納する場所を予約
    Dim rng As Range, r As Range
    '受注シート内の最終セルまでを予約した場所に格納
    Set rng = Range(wsh.Range("A2"), wsh.Cells.SpecialCells(xlCellTypeLastCell))
    
    '受注シート内の表を下から1行ずつ取り出して処理
    Dim i As Long
    For i = rng.Rows.Count To 1 Step -1
        
        '受注数が空欄でない場合
        If Not IsEmpty(rng.Cells(i, 3)) Then
            '何もしない
        
        '受注数が空欄かつ備考欄に「削除」または「不要」の文字が含まれている場合
        ElseIf rng.Cells(i, 4).Text Like "*削除*" _
            Or rng.Cells(i, 4).Text Like "*不要*" Then
            '行全体を削除
            rng.Cells(i, 4).EntireRow.Delete
        
        Else
            '何もしない
        End If
    
    Next i
        
    '見た目を整える
    wsh.Activate
    wsh.Range("A1").Activate
    
    '== 余談 ==
    
    '・備考欄の判定を Select Case にしようか迷いましたが、
    '　そうすると Delete が2箇所に発生するためやめました。
    '　というかそもそも Like が使えないのね。
    
    '・ForEach で上の行から操作するとアジャパーなので
    '　ForNext で下の行から処理するようにしました。
    '　上の行から消していくと、2行連続で削除対象の場合にアジャパーします。
    
End Sub
