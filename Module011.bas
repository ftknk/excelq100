Attribute VB_Name = "Module011"
Option Explicit

Sub Q011()
    
    '今表示されているシートを確保
    Dim wsh As Worksheet
    Set wsh = ActiveSheet
    
    '使われている範囲を確保
    Dim rng As Range
    Set rng = wsh.UsedRange
    
    '使われている範囲のセルを1つずつ走査
    Dim r As Range
    For Each r In rng
        If Not r.MergeCells Then
            '結合セルでない場合、何もしない
        ElseIf r.MergeArea.Cells(1).Address = r.Address Then
            '結合セルの左上だった場合、メモを設定
            r.NoteText "警告：結合されたセル"
            'メモのサイズを文字にあわせる
            r.Comment.Shape.TextFrame.AutoSize = True
            'メモを表示したままにする
            r.Comment.Visible = True
        Else
            '結合セルの左上以外の場合、何もしない
            '結合されているのでどれか1つにメモをつければよい。
        End If
    Next r
        
    '見た目を整える
    wsh.Activate
    wsh.Range("A1").Activate

End Sub
Sub Q012()
    
    'シートを確保
    Dim wsh As Worksheet
    Set wsh = Sheet012
    
    '使われている範囲を確保
    Dim rng As Range
    Set rng = wsh.UsedRange
    
    '金額の端数、セルの数
    Dim c As Currency, i As Long
    
    '使われている範囲のC列のセルを1つずつ走査
    Dim r As Range
    For Each r In rng.Columns(3).Cells
        If r.MergeCells Then
            c = 0
            
            '結合セルだった場合、セル結合を解除し、入っている金額を整数で均等に割り振る
            With r.MergeArea
                
                'セルの数を取得
                i = .Cells.Count
                '金額の端数を取得（小数点以下切り捨て）
                c = r.Value Mod i
                
                'セル結合を解除
                .UnMerge
                
                '分割した金額を設定
                .Value = (r.Value - c) / i
            
            End With
            
            '端数がある場合、上から端数分 +1 する
            If c > 0 Then r.Resize(c, 1).Value = r.Value + 1
        
        Else
            '結合セル以外の場合、何もしない
        End If
    Next r
        
    '見た目を整える
    wsh.Activate
    wsh.Range("A1").Activate

End Sub
Sub Q013()
    
    'セル以外（図形等）が選択されている場合は何もせずに正常終了する
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    '検索する文字を設定
    Const WHAT As String = "注意"
    Dim LNG As Long: LNG = Len(WHAT)
    
    '選択範囲の文字が設定されているセルを確保
    Dim rng As Range
    Set rng = Selection
    If rng.Cells.CountLarge > 1 Then
        'たくさん選択されていた場合、文字列が入っているセルだけに絞る
        On Error Resume Next
            Set rng = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
            '文字列が無ければ処理終了
            If Err.Number > 0 Then Exit Sub
        On Error GoTo 0
    End If
    
    '確保したセルを1つずつ走査
    Dim r As Range
    For Each r In rng.Cells
        
        '文字列に「注意」という文字があった場合
        If r.Value Like "*" & WHAT & "*" Then
            
            'すべての「注意」を探す
            Dim i As Long: i = 0
            Do
                '「注意」が見つかった場合
                If i > 0 Then
                    '「注意」の文字だけを"赤の太字"に設定
                    With r.Characters(i, LNG).Font
                        .Color = vbRed
                        .Bold = True
                    End With
                End If
                
                '次の「注意」を検索
                i = i + 1
                i = InStr(i, r.Value, WHAT, vbTextCompare)
            
            Loop While i > 0
        
        Else
            '文字列に「注意」という文字がない場合、何もしない
        End If
    Next r

    '== 余談 ==
    '・Loop文苦手やねん…
    '・「注意」が見つかった場合…のIF文は捨てなのですが、
    '　InStrを2回書くよりもこっちの方が好きかなぁ。

End Sub
Sub Q014()
    
    'ブックを確保
    Dim wbk As Workbook
    Set wbk = ActiveWorkbook
    
    '一応確認
    Dim msgResult As VbMsgBoxResult
    msgResult = MsgBox("実行前にファイルを保存しますか？" & vbCrLf _
        & "「はい」保存して実行" & vbCrLf _
        & "「いいえ」保存せず実行" & vbCrLf _
        & "「キャンセル」実行せず戻る", vbYesNoCancel + vbInformation, "事前確認")
    Select Case msgResult
    Case vbYes: wbk.Save
    Case vbNo: '何もしない
    Case Else: Exit Sub
    End Select

    '結果表示用のシート作成
    Dim wshResult As Worksheet
    Workbooks.Add
    Set wshResult = ActiveSheet
    
    '体裁を整える（やさしい）
    wshResult.Range("A1").ColumnWidth = 32
    wshResult.Range("B1").ColumnWidth = 80
    
    '注意書き（やさしい）
    wshResult.Range("A1").Value = "＜客先提出ファイルの注意事項＞"
    wshResult.Range("A2").Value = "・上書きせず、別名で保存してください。"
    wshResult.Range("A3").Value = "・ファイルサイズに注意してください。4Mを超える場合はサイズを減らす努力をしてください。"
    wshResult.Range("A4").Value = "・アクティブセル、表示形式、倍率などは適宜調整願います。"
    
    '見出し（やさしい）
    wshResult.Range("A6").Value = "＜処理結果＞　処理日時：" _
        & Format(Now, "yyyy/mm/dd hh:mm:ss") _
        & "、処理前保存：" & IIf(msgResult = vbYes, "済", "不明")
    wshResult.Range("A7").Value = "-- シート名 ----"
    wshResult.Range("B7").Value = "-- 結果 --------"
    
    '結果を入れる行位置
    Dim i As Long: i = 8
    
    'シート単位でループ
    Dim wsh As Worksheet
    For Each wsh In wbk.Worksheets
    
        'シート名
        wshResult.Cells(i, 1).Value = wsh.Name
    
        '非表示の場合、表示する（問答無用）
        If Not wsh.Visible = xlSheetVisible Then
            wsh.Visible = xlSheetVisible
            wshResult.Cells(i, 2).Value = wshResult.Cells(i, 2).Value & "・非表示に設定されていたため、再表示しました。" & vbCrLf
        End If
    
        'シート名に「社外秘」の文字が含まれるシートを削除（問答無用）
        If wsh.Name Like "*社外秘*" Then
            Application.DisplayAlerts = False
            wsh.Delete
            Application.DisplayAlerts = True
            wshResult.Cells(i, 2).Value = wshResult.Cells(i, 2).Value & "・「社外秘」が含まれていたため、削除しました。" & vbCrLf
        Else
        
            'シート名に「社外秘」の文字が含まれない場合
            '計算式を消して値だけにする
            Dim errNo As Long
            On Error Resume Next
                wsh.Cells.SpecialCells xlCellTypeFormulas
                errNo = Err.Number
                Err.Clear
            On Error GoTo 0
            If errNo = 0 Then
                If Not wsh.EnableCalculation Then
                    wshResult.Cells(i, 2).Value = wshResult.Cells(i, 2).Value & "・自動計算がオフになっているため、数式の妥当性は保証されません。" & vbCrLf
                End If
                'セルをコピーして、値で貼り付け（問答無用）
                wsh.Cells.Copy
                wsh.Cells.PasteSpecial xlPasteValues
                Application.CutCopyMode = False
                wshResult.Cells(i, 2).Value = wshResult.Cells(i, 2).Value & "・数式を値に変更しました。" & vbCrLf
            End If
        
        End If
        
        If IsEmpty(wshResult.Cells(i, 2).Value) Then
            '結果に値が入っていない場合、対応なしと表示
            wshResult.Cells(i, 2).Value = "（対応なし）"
        Else
            '結果に値が入っている場合、末尾の改行を削る
            wshResult.Cells(i, 2).Value = Left(wshResult.Cells(i, 2).Value, Len(wshResult.Cells(i, 2).Value) - 2)
        End If
                
        i = i + 1
    Next wsh
    
    '== 余談 ==
    '・結果表示用のブックは、余分なシートがあれば消すと親切。
    '・注意書きの内容も実装できなくはないけれども、ツールに任せっぱは良くないからね。
    '・単純に「wbk.Save」すると、未保存のブックは「どこか」に保存されます。
    '　（その時にデフォルトのパスに保存されるので、下手するとファイルサーバとかね）

End Sub
Sub Q015_01()

    'よいこ
    With ActiveWorkbook
        .Worksheets("2020年04月").Move .Worksheets(1)
        .Worksheets("2020年05月").Move .Worksheets(2)
        .Worksheets("2020年06月").Move .Worksheets(3)
        .Worksheets("2020年07月").Move .Worksheets(4)
        .Worksheets("2020年08月").Move .Worksheets(5)
        .Worksheets("2020年09月").Move .Worksheets(6)
        .Worksheets("2020年10月").Move .Worksheets(7)
        .Worksheets("2020年11月").Move .Worksheets(8)
        .Worksheets("2020年12月").Move .Worksheets(9)
        .Worksheets("2021年01月").Move .Worksheets(10)
        .Worksheets("2021年02月").Move .Worksheets(11)
        .Worksheets("2021年03月").Move .Worksheets(12)
    End With

End Sub
Sub Q015_02()

    'わるいこ
    Dim i As Integer
    For i = 1 To 12
        ActiveWorkbook.Worksheets(Format(DateAdd("m", i, "2020/3/1"), "yyyy年mm月")).Move ActiveWorkbook.Worksheets(i)
    Next i

End Sub
Sub Q015_03()

    'ふつうのこ
    Dim wshA As Worksheet
    For Each wshA In ActiveWorkbook.Worksheets
        Dim wshB As Worksheet
        For Each wshB In ActiveWorkbook.Worksheets
            If StrComp(wshA.Name, wshB.Name) = -1 Then
                 wshA.Move wshB
                 Exit For
            End If
        Next wshB
    Next wshA

End Sub
Sub Q016()

    '選択範囲に対して実施（個人の趣味）
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rngSelection As Range: Set rngSelection = Selection

    '選択されたセルをループ
    Dim rng As Range
    For Each rng In rngSelection
        'セルの内容判定
        If IsEmpty(rng) Then
            '空の場合何もしない
        ElseIf rng.HasFormula Then
            '数式の場合何もしない
        Else
            '文字列の場合、無駄な改行を削除
            Dim s As String: s = rng.Value

            'CRLFはLFに変換する
            s = Replace(s, vbCrLf, vbLf)
    
            'LFが連続している場合、1つにする
            Do While s Like "*" & vbLf & vbLf & "*"
                s = Replace(s, vbLf & vbLf, vbLf)
            Loop
            
            '先頭の改行を削除
            If Left(s, 1) = vbLf Then s = Mid(s, 2)
            
            '末尾の改行を削除
            If Right(s, 1) = vbLf Then s = Left(s, Len(s) - 1)
            
            '変換した結果を設定
            rng.Value = s
        End If
    Next rng
    
    '== 余談 ==
    '・「Mid(s, 2)」と「Right(s,Len(s) - 1)」どっちが軽いだろう？
    '・最後に結果をセルに戻すの忘れがち(´･ω･`)

End Sub
Sub Q017_11()
    
    'RemoveDuplicates
    Sheet017_1.Range("C1").Resize( _
        Sheet017_1.Range("A1").CurrentRegion.Rows.Count, 4).Copy
    Sheet017_2.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Sheet017_2.Range("A1").CurrentRegion.RemoveDuplicates _
        Columns:=Array(1, 2, 3, 4), Header:=xlYes

End Sub
Sub Q017_12()
    
    'Scripting.Dictionary
    Dim var As Variant
    var = Sheet017_1.Range("C1").Resize( _
        Sheet017_1.Range("A1").CurrentRegion.Rows.Count, 4)
    Dim v As Variant, i As Long, s As String
    Dim dic As Dictionary: Set dic = New Dictionary
    For i = LBound(var) To UBound(var)
        s = Join(Array(var(i, 1), var(i, 2), var(i, 3), var(i, 4)), Chr(0))
        dic.Item(s) = i
    Next i
    var = dic.Keys
    Dim rng As Range: Set rng = Sheet017_2.Range("A1:D1")
    For Each v In var
        rng = Split(v, Chr(0))
        Set rng = rng.Offset(1)
    Next v

End Sub

Sub Q018()
    
    '今表示されているブックを確保
    Dim wbk As Workbook
    Set wbk = ActiveWorkbook
    
    '各件数
    Dim hiddenCount As Long: hiddenCount = 0
    Dim deleteCount As Long: deleteCount = 0
    Dim errorCount As Long: errorCount = 0
    'wbk.Names(2).Visible = False
    'wbk.Names(3).Visible = False

    'ブックが保持する名前をすべて走査
    Dim nm As Name
    For Each nm In wbk.Names
        Dim r As Range
On Error Resume Next
            Set r = nm.RefersToRange
On Error GoTo Err_Nm
        If r Is Nothing Then
            '範囲を参照できない場合
            'イミディエイトに「名前」と「参照範囲」を出力
            Debug.Print nm.Name & ", " & nm.RefersTo
            '名前を削除
            nm.Delete
            deleteCount = deleteCount + 1
        ElseIf Not nm.Visible Then
            '名前定義が非表示の場合、表示に変更
            nm.Visible = True
            hiddenCount = hiddenCount + 1
            'Debug.Print "◎" & nm.Name & ", " & nm.RefersTo
        Else
            'Debug.Print "○" & nm.Name & ", " & nm.RefersTo
        End If
        GoTo End_Nm
Err_Nm:
    errorCount = errorCount + 1
    Debug.Print "-- ↑削除失敗 --------"
    Debug.Print Err.Description
    Debug.Print "----------------------"
    Err.Clear
End_Nm:
    On Error GoTo 0
    Next
    
    '非表示件数と削除件数をメッセージボックスに表示
    MsgBox "名前定義の整理が完了しました。" & _
        vbCrLf & "　削除：" & deleteCount & "件" & _
        vbCrLf & "　表示：" & hiddenCount & "件" _
        , vbOKOnly + vbInformation, "完了"
    If errorCount > 0 Then
        MsgBox "削除できなかった名前定義があります。" & _
            vbCrLf & "確認してください。" & _
            vbCrLf & "　失敗：" & errorCount & "件" _
            , vbOKOnly + vbExclamation, "警告" _
            , "https://support.office.com/client/results", 25450
    End If
    
    '== 余談 ==
    '・「非表示」かつ「エラー」のパターンもあるので、削除を先にしました。
    '・本当は、ダイアログのヘルプボタンを活用したかった。
    '　↓残骸
    '　vbMsgBoxHelpButton, "https://support.office.com/client/results", 25450

End Sub

Sub Q019()
    Call Q019_sub(ActiveSheet)
End Sub
Sub Q019_sub(ByRef wsh As Worksheet)
Dim shp As Shape
For Each shp In wsh.Shapes
    
    '名前に「Q019」が含まれている場合、コピー後の図形のためスルー
    If Left(shp.Name, 5) = "Q019_" Then GoTo Next_Shp
    
    'Excel的にオートシェイプ以外の場合、スルー
    '（線やコネクタ、スマートアート、グラフ、フォームなど）
    If shp.AutoShapeType = msoShapeMixed Then GoTo Next_Shp
    
    '画像をコピー＆ペースト
    With shp.Duplicate
        '増殖防止のため名前を変更
        .Name = "Q019_cpy_" & shp.Name
        shp.Name = "Q019_org_" & shp.Name
        '画像の位置をコピー元の真横に移動
        .Top = shp.Top
        .Left = shp.Left + shp.Width
    End With

Next_Shp:
Next

    '== 余談 ==
    '・shp.Type の方が正確だが、パターン網羅した実装がめんどかった。
    '　まずはサクッと実装して、要望があればそれに応えていくスタイル。
    '・「無限増殖防止」についても、コピー元は変えないであればそうする。

End Sub

Sub Q020()

    'バックアップフォルダのパスを作成
    Dim bkDir As String: bkDir = ThisWorkbook.Path & "\BACKUP\"

    'マクロブック(ThisWorkbook)と同じフォルダに"BACKUP"フォルダを作成
    If Dir(bkDir, vbDirectory) = "" Then MkDir bkDir
    
    'バックアップファイル名を作成
    Dim bkFile As String
    bkFile = Replace(ThisWorkbook.Name, ".xlsm", "") & Format(Now(), "_yyyymmddhhmm") & ".xlsm"
    
    '同じファイルがあった場合、上書き確認（キャンセルの場合は処理終了）
    If Dir(bkDir & bkFile) <> "" Then
        If vbCancel = _
            MsgBox("直前にバックアップしています。" & vbCrLf & "上書きしてよろしいですか？" _
            & vbCrLf & bkFile, _
                vbOKCancel + vbCritical, "確認") Then Exit Sub
    End If

    '今の状態をバックアップ
    ThisWorkbook.SaveCopyAs bkDir & bkFile

End Sub
