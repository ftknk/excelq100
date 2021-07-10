Attribute VB_Name = "Module011"
Option Explicit

Sub Q011()
    
    '���\������Ă���V�[�g���m��
    Dim wsh As Worksheet
    Set wsh = ActiveSheet
    
    '�g���Ă���͈͂��m��
    Dim rng As Range
    Set rng = wsh.UsedRange
    
    '�g���Ă���͈͂̃Z����1������
    Dim r As Range
    For Each r In rng
        If Not r.MergeCells Then
            '�����Z���łȂ��ꍇ�A�������Ȃ�
        ElseIf r.MergeArea.Cells(1).Address = r.Address Then
            '�����Z���̍��ゾ�����ꍇ�A������ݒ�
            r.NoteText "�x���F�������ꂽ�Z��"
            '�����̃T�C�Y�𕶎��ɂ��킹��
            r.Comment.Shape.TextFrame.AutoSize = True
            '������\�������܂܂ɂ���
            r.Comment.Visible = True
        Else
            '�����Z���̍���ȊO�̏ꍇ�A�������Ȃ�
            '��������Ă���̂łǂꂩ1�Ƀ���������΂悢�B
        End If
    Next r
        
    '�����ڂ𐮂���
    wsh.Activate
    wsh.Range("A1").Activate

End Sub
Sub Q012()
    
    '�V�[�g���m��
    Dim wsh As Worksheet
    Set wsh = Sheet012
    
    '�g���Ă���͈͂��m��
    Dim rng As Range
    Set rng = wsh.UsedRange
    
    '���z�̒[���A�Z���̐�
    Dim c As Currency, i As Long
    
    '�g���Ă���͈͂�C��̃Z����1������
    Dim r As Range
    For Each r In rng.Columns(3).Cells
        If r.MergeCells Then
            c = 0
            
            '�����Z���������ꍇ�A�Z���������������A�����Ă�����z�𐮐��ŋϓ��Ɋ���U��
            With r.MergeArea
                
                '�Z���̐����擾
                i = .Cells.Count
                '���z�̒[�����擾�i�����_�ȉ��؂�̂āj
                c = r.Value Mod i
                
                '�Z������������
                .UnMerge
                
                '�����������z��ݒ�
                .Value = (r.Value - c) / i
            
            End With
            
            '�[��������ꍇ�A�ォ��[���� +1 ����
            If c > 0 Then r.Resize(c, 1).Value = r.Value + 1
        
        Else
            '�����Z���ȊO�̏ꍇ�A�������Ȃ�
        End If
    Next r
        
    '�����ڂ𐮂���
    wsh.Activate
    wsh.Range("A1").Activate

End Sub
Sub Q013()
    
    '�Z���ȊO�i�}�`���j���I������Ă���ꍇ�͉��������ɐ���I������
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    '�������镶����ݒ�
    Const WHAT As String = "����"
    Dim LNG As Long: LNG = Len(WHAT)
    
    '�I��͈͂̕������ݒ肳��Ă���Z�����m��
    Dim rng As Range
    Set rng = Selection
    If rng.Cells.CountLarge > 1 Then
        '��������I������Ă����ꍇ�A�����񂪓����Ă���Z�������ɍi��
        On Error Resume Next
            Set rng = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
            '�����񂪖�����Ώ����I��
            If Err.Number > 0 Then Exit Sub
        On Error GoTo 0
    End If
    
    '�m�ۂ����Z����1������
    Dim r As Range
    For Each r In rng.Cells
        
        '������Ɂu���Ӂv�Ƃ����������������ꍇ
        If r.Value Like "*" & WHAT & "*" Then
            
            '���ׂẮu���Ӂv��T��
            Dim i As Long: i = 0
            Do
                '�u���Ӂv�����������ꍇ
                If i > 0 Then
                    '�u���Ӂv�̕���������"�Ԃ̑���"�ɐݒ�
                    With r.Characters(i, LNG).Font
                        .Color = vbRed
                        .Bold = True
                    End With
                End If
                
                '���́u���Ӂv������
                i = i + 1
                i = InStr(i, r.Value, WHAT, vbTextCompare)
            
            Loop While i > 0
        
        Else
            '������Ɂu���Ӂv�Ƃ����������Ȃ��ꍇ�A�������Ȃ�
        End If
    Next r

    '== �]�k ==
    '�ELoop������˂�c
    '�E�u���Ӂv�����������ꍇ�c��IF���͎̂ĂȂ̂ł����A
    '�@InStr��2�񏑂������������̕����D�����Ȃ��B

End Sub
Sub Q014()
    
    '�u�b�N���m��
    Dim wbk As Workbook
    Set wbk = ActiveWorkbook
    
    '�ꉞ�m�F
    Dim msgResult As VbMsgBoxResult
    msgResult = MsgBox("���s�O�Ƀt�@�C����ۑ����܂����H" & vbCrLf _
        & "�u�͂��v�ۑ����Ď��s" & vbCrLf _
        & "�u�������v�ۑ��������s" & vbCrLf _
        & "�u�L�����Z���v���s�����߂�", vbYesNoCancel + vbInformation, "���O�m�F")
    Select Case msgResult
    Case vbYes: wbk.Save
    Case vbNo: '�������Ȃ�
    Case Else: Exit Sub
    End Select

    '���ʕ\���p�̃V�[�g�쐬
    Dim wshResult As Worksheet
    Workbooks.Add
    Set wshResult = ActiveSheet
    
    '�̍ق𐮂���i�₳�����j
    wshResult.Range("A1").ColumnWidth = 32
    wshResult.Range("B1").ColumnWidth = 80
    
    '���ӏ����i�₳�����j
    wshResult.Range("A1").Value = "���q���o�t�@�C���̒��ӎ�����"
    wshResult.Range("A2").Value = "�E�㏑�������A�ʖ��ŕۑ����Ă��������B"
    wshResult.Range("A3").Value = "�E�t�@�C���T�C�Y�ɒ��ӂ��Ă��������B4M�𒴂���ꍇ�̓T�C�Y�����炷�w�͂����Ă��������B"
    wshResult.Range("A4").Value = "�E�A�N�e�B�u�Z���A�\���`���A�{���Ȃǂ͓K�X�����肢�܂��B"
    
    '���o���i�₳�����j
    wshResult.Range("A6").Value = "���������ʁ��@���������F" _
        & Format(Now, "yyyy/mm/dd hh:mm:ss") _
        & "�A�����O�ۑ��F" & IIf(msgResult = vbYes, "��", "�s��")
    wshResult.Range("A7").Value = "-- �V�[�g�� ----"
    wshResult.Range("B7").Value = "-- ���� --------"
    
    '���ʂ�����s�ʒu
    Dim i As Long: i = 8
    
    '�V�[�g�P�ʂŃ��[�v
    Dim wsh As Worksheet
    For Each wsh In wbk.Worksheets
    
        '�V�[�g��
        wshResult.Cells(i, 1).Value = wsh.Name
    
        '��\���̏ꍇ�A�\������i�ⓚ���p�j
        If Not wsh.Visible = xlSheetVisible Then
            wsh.Visible = xlSheetVisible
            wshResult.Cells(i, 2).Value = wshResult.Cells(i, 2).Value & "�E��\���ɐݒ肳��Ă������߁A�ĕ\�����܂����B" & vbCrLf
        End If
    
        '�V�[�g���Ɂu�ЊO��v�̕������܂܂��V�[�g���폜�i�ⓚ���p�j
        If wsh.Name Like "*�ЊO��*" Then
            Application.DisplayAlerts = False
            wsh.Delete
            Application.DisplayAlerts = True
            wshResult.Cells(i, 2).Value = wshResult.Cells(i, 2).Value & "�E�u�ЊO��v���܂܂�Ă������߁A�폜���܂����B" & vbCrLf
        Else
        
            '�V�[�g���Ɂu�ЊO��v�̕������܂܂�Ȃ��ꍇ
            '�v�Z���������Ēl�����ɂ���
            Dim errNo As Long
            On Error Resume Next
                wsh.Cells.SpecialCells xlCellTypeFormulas
                errNo = Err.Number
                Err.Clear
            On Error GoTo 0
            If errNo = 0 Then
                If Not wsh.EnableCalculation Then
                    wshResult.Cells(i, 2).Value = wshResult.Cells(i, 2).Value & "�E�����v�Z���I�t�ɂȂ��Ă��邽�߁A�����̑Ó����͕ۏ؂���܂���B" & vbCrLf
                End If
                '�Z�����R�s�[���āA�l�œ\��t���i�ⓚ���p�j
                wsh.Cells.Copy
                wsh.Cells.PasteSpecial xlPasteValues
                Application.CutCopyMode = False
                wshResult.Cells(i, 2).Value = wshResult.Cells(i, 2).Value & "�E������l�ɕύX���܂����B" & vbCrLf
            End If
        
        End If
        
        If IsEmpty(wshResult.Cells(i, 2).Value) Then
            '���ʂɒl�������Ă��Ȃ��ꍇ�A�Ή��Ȃ��ƕ\��
            wshResult.Cells(i, 2).Value = "�i�Ή��Ȃ��j"
        Else
            '���ʂɒl�������Ă���ꍇ�A�����̉��s�����
            wshResult.Cells(i, 2).Value = Left(wshResult.Cells(i, 2).Value, Len(wshResult.Cells(i, 2).Value) - 2)
        End If
                
        i = i + 1
    Next wsh
    
    '== �]�k ==
    '�E���ʕ\���p�̃u�b�N�́A�]���ȃV�[�g������Ώ����Ɛe�؁B
    '�E���ӏ����̓��e�������ł��Ȃ��͂Ȃ�����ǂ��A�c�[���ɔC�����ς͗ǂ��Ȃ�����ˁB
    '�E�P���Ɂuwbk.Save�v����ƁA���ۑ��̃u�b�N�́u�ǂ����v�ɕۑ�����܂��B
    '�@�i���̎��Ƀf�t�H���g�̃p�X�ɕۑ������̂ŁA���肷��ƃt�@�C���T�[�o�Ƃ��ˁj

End Sub
Sub Q015_01()

    '�悢��
    With ActiveWorkbook
        .Worksheets("2020�N04��").Move .Worksheets(1)
        .Worksheets("2020�N05��").Move .Worksheets(2)
        .Worksheets("2020�N06��").Move .Worksheets(3)
        .Worksheets("2020�N07��").Move .Worksheets(4)
        .Worksheets("2020�N08��").Move .Worksheets(5)
        .Worksheets("2020�N09��").Move .Worksheets(6)
        .Worksheets("2020�N10��").Move .Worksheets(7)
        .Worksheets("2020�N11��").Move .Worksheets(8)
        .Worksheets("2020�N12��").Move .Worksheets(9)
        .Worksheets("2021�N01��").Move .Worksheets(10)
        .Worksheets("2021�N02��").Move .Worksheets(11)
        .Worksheets("2021�N03��").Move .Worksheets(12)
    End With

End Sub
Sub Q015_02()

    '��邢��
    Dim i As Integer
    For i = 1 To 12
        ActiveWorkbook.Worksheets(Format(DateAdd("m", i, "2020/3/1"), "yyyy�Nmm��")).Move ActiveWorkbook.Worksheets(i)
    Next i

End Sub
Sub Q015_03()

    '�ӂ��̂�
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

    '�I��͈͂ɑ΂��Ď��{�i�l�̎�j
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rngSelection As Range: Set rngSelection = Selection

    '�I�����ꂽ�Z�������[�v
    Dim rng As Range
    For Each rng In rngSelection
        '�Z���̓��e����
        If IsEmpty(rng) Then
            '��̏ꍇ�������Ȃ�
        ElseIf rng.HasFormula Then
            '�����̏ꍇ�������Ȃ�
        Else
            '������̏ꍇ�A���ʂȉ��s���폜
            Dim s As String: s = rng.Value

            'CRLF��LF�ɕϊ�����
            s = Replace(s, vbCrLf, vbLf)
    
            'LF���A�����Ă���ꍇ�A1�ɂ���
            Do While s Like "*" & vbLf & vbLf & "*"
                s = Replace(s, vbLf & vbLf, vbLf)
            Loop
            
            '�擪�̉��s���폜
            If Left(s, 1) = vbLf Then s = Mid(s, 2)
            
            '�����̉��s���폜
            If Right(s, 1) = vbLf Then s = Left(s, Len(s) - 1)
            
            '�ϊ��������ʂ�ݒ�
            rng.Value = s
        End If
    Next rng
    
    '== �]�k ==
    '�E�uMid(s, 2)�v�ƁuRight(s,Len(s) - 1)�v�ǂ������y�����낤�H
    '�E�Ō�Ɍ��ʂ��Z���ɖ߂��̖Y�ꂪ��(�L��֥`)

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
    
    '���\������Ă���u�b�N���m��
    Dim wbk As Workbook
    Set wbk = ActiveWorkbook
    
    '�e����
    Dim hiddenCount As Long: hiddenCount = 0
    Dim deleteCount As Long: deleteCount = 0
    Dim errorCount As Long: errorCount = 0
    'wbk.Names(2).Visible = False
    'wbk.Names(3).Visible = False

    '�u�b�N���ێ����閼�O�����ׂđ���
    Dim nm As Name
    For Each nm In wbk.Names
        Dim r As Range
On Error Resume Next
            Set r = nm.RefersToRange
On Error GoTo Err_Nm
        If r Is Nothing Then
            '�͈͂��Q�Ƃł��Ȃ��ꍇ
            '�C�~�f�B�G�C�g�Ɂu���O�v�Ɓu�Q�Ɣ͈́v���o��
            Debug.Print nm.Name & ", " & nm.RefersTo
            '���O���폜
            nm.Delete
            deleteCount = deleteCount + 1
        ElseIf Not nm.Visible Then
            '���O��`����\���̏ꍇ�A�\���ɕύX
            nm.Visible = True
            hiddenCount = hiddenCount + 1
            'Debug.Print "��" & nm.Name & ", " & nm.RefersTo
        Else
            'Debug.Print "��" & nm.Name & ", " & nm.RefersTo
        End If
        GoTo End_Nm
Err_Nm:
    errorCount = errorCount + 1
    Debug.Print "-- ���폜���s --------"
    Debug.Print Err.Description
    Debug.Print "----------------------"
    Err.Clear
End_Nm:
    On Error GoTo 0
    Next
    
    '��\�������ƍ폜���������b�Z�[�W�{�b�N�X�ɕ\��
    MsgBox "���O��`�̐������������܂����B" & _
        vbCrLf & "�@�폜�F" & deleteCount & "��" & _
        vbCrLf & "�@�\���F" & hiddenCount & "��" _
        , vbOKOnly + vbInformation, "����"
    If errorCount > 0 Then
        MsgBox "�폜�ł��Ȃ��������O��`������܂��B" & _
            vbCrLf & "�m�F���Ă��������B" & _
            vbCrLf & "�@���s�F" & errorCount & "��" _
            , vbOKOnly + vbExclamation, "�x��" _
            , "https://support.office.com/client/results", 25450
    End If
    
    '== �]�k ==
    '�E�u��\���v���u�G���[�v�̃p�^�[��������̂ŁA�폜���ɂ��܂����B
    '�E�{���́A�_�C�A���O�̃w���v�{�^�������p�����������B
    '�@���c�[
    '�@vbMsgBoxHelpButton, "https://support.office.com/client/results", 25450

End Sub

Sub Q019()
    Call Q019_sub(ActiveSheet)
End Sub
Sub Q019_sub(ByRef wsh As Worksheet)
Dim shp As Shape
For Each shp In wsh.Shapes
    
    '���O�ɁuQ019�v���܂܂�Ă���ꍇ�A�R�s�[��̐}�`�̂��߃X���[
    If Left(shp.Name, 5) = "Q019_" Then GoTo Next_Shp
    
    'Excel�I�ɃI�[�g�V�F�C�v�ȊO�̏ꍇ�A�X���[
    '�i����R�l�N�^�A�X�}�[�g�A�[�g�A�O���t�A�t�H�[���Ȃǁj
    If shp.AutoShapeType = msoShapeMixed Then GoTo Next_Shp
    
    '�摜���R�s�[���y�[�X�g
    With shp.Duplicate
        '���B�h�~�̂��ߖ��O��ύX
        .Name = "Q019_cpy_" & shp.Name
        shp.Name = "Q019_org_" & shp.Name
        '�摜�̈ʒu���R�s�[���̐^���Ɉړ�
        .Top = shp.Top
        .Left = shp.Left + shp.Width
    End With

Next_Shp:
Next

    '== �]�k ==
    '�Eshp.Type �̕������m�����A�p�^�[���ԗ������������߂�ǂ������B
    '�@�܂��̓T�N�b�Ǝ������āA�v�]������΂���ɉ����Ă����X�^�C���B
    '�E�u�������B�h�~�v�ɂ��Ă��A�R�s�[���͕ς��Ȃ��ł���΂�������B

End Sub

Sub Q020()

    '�o�b�N�A�b�v�t�H���_�̃p�X���쐬
    Dim bkDir As String: bkDir = ThisWorkbook.Path & "\BACKUP\"

    '�}�N���u�b�N(ThisWorkbook)�Ɠ����t�H���_��"BACKUP"�t�H���_���쐬
    If Dir(bkDir, vbDirectory) = "" Then MkDir bkDir
    
    '�o�b�N�A�b�v�t�@�C�������쐬
    Dim bkFile As String
    bkFile = Replace(ThisWorkbook.Name, ".xlsm", "") & Format(Now(), "_yyyymmddhhmm") & ".xlsm"
    
    '�����t�@�C�����������ꍇ�A�㏑���m�F�i�L�����Z���̏ꍇ�͏����I���j
    If Dir(bkDir & bkFile) <> "" Then
        If vbCancel = _
            MsgBox("���O�Ƀo�b�N�A�b�v���Ă��܂��B" & vbCrLf & "�㏑�����Ă�낵���ł����H" _
            & vbCrLf & bkFile, _
                vbOKCancel + vbCritical, "�m�F") Then Exit Sub
    End If

    '���̏�Ԃ��o�b�N�A�b�v
    ThisWorkbook.SaveCopyAs bkDir & bkFile

End Sub
