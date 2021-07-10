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
    '�\�ȊO�̏�񂪓����Ă�����A�E�g�ł�����ǂ��c
    Range(Range("b2"), Cells.SpecialCells(xlCellTypeLastCell)).ClearContents
End Sub
Sub Q004()
    '�ŏI�Z���܂Ł��萔�Z����I���i�\�ȊO�̏�񂪓����Ă�����A�E�g�j
    Range(Range("b2"), Cells.SpecialCells(xlCellTypeLastCell)) _
        .SpecialCells(xlCellTypeConstants).ClearContents
End Sub
Sub Q005()
    '�\�͈̔͂ƍs���i�[����ꏊ��\��
    Dim rng As Range, r As Range
    '�ŏI�Z���܂ł�\�񂵂��ꏊ�Ɋi�[�iD��Ɍ��o���������Ă���O��j
    Set rng = Range(Range("B3"), Cells.SpecialCells(xlCellTypeLastCell))
    
    '�\��1�s�����o���ď���
    For Each r In rng.Rows '�@
        
        'B��܂���C�񂪋󗓂��ǂ����m�F
        If IsEmpty(r.Cells(1)) Or IsEmpty(r.Cells(2)) Then '�B
            '�ǂ��炩���󗓂������牽�����Ȃ� �C
        
        ElseIf IsNumeric(r.Cells(1)) And IsNumeric(r.Cells(2)) Then
            '���ł����琔�l�`�F�b�N�����Ă���
            '�����Ƃ����l�������Ă����琔�������A������ݒ肷��
            r.Cells(3).Formula = "=B:B*C:C" '�����uB��~C��v
            r.Cells(3).NumberFormatLocal = "\#,##0" '�����u�ʉ�\�̃J���}�ҏW�v

        Else
            '�����ɗ�����A�ǂ��炩�����l�ł͖��������Ƃ������Ƃł�
        
        End If
    
    Next r '�A
    
    '== �]�k ==
    '�@ For Each r ... rng.Rows �� rng ����1�s���o���� r �Ɋi�[
    '�A�uNext�v�ŏI��点���ɁuNext r�v�ɂ���ƈ��S�B
    '�B r.Cells(n) �́A���ォ��E���ւƐ����ĉ��ڂ̃Z�����������B
    '�C if�� �́u�������Ȃ��v�͐��i�ł��Ɠ{���邯��ǂ��A
    '�@ �݌v���̖����Г������֗��c�[���͓ǂ݂₷���悤��
    '�@ if �� else ���������ăR�����g�����悤�ɂ��Ă���B
End Sub
Sub Q006()
    '�\�͈̔͂ƍs���i�[����ꏊ��\��
    Dim rng As Range, r As Range
    '�ŏI�Z���܂ł�\�񂵂��ꏊ�Ɋi�[�iD��Ɍ��o���������Ă���O��j
    Set rng = Range(Range("A2"), Cells.SpecialCells(xlCellTypeLastCell))
    
    '�\��1�s�����o���ď���
    For Each r In rng.Rows '�@
        
        'B��܂���C�񂪋󗓂��ǂ����m�F
        If r.Cells(1).Text Like "*-*" Then
            '���i�R�[�h��"-"���܂܂�Ă����牽�����Ȃ� �D
        
        ElseIf IsEmpty(r.Cells(2)) Or IsEmpty(r.Cells(3)) Then '�B
            '�ǂ��炩���󗓂������牽�����Ȃ� �C
        
        ElseIf IsNumeric(r.Cells(2)) And IsNumeric(r.Cells(3)) Then
            '���ł����琔�l�`�F�b�N�����Ă���
            '�����Ƃ����l�������Ă����琔�������A������ݒ肷��
            r.Cells(4).Formula = "=B:B*C:C" '�����uB��~C��v
            r.Cells(4).NumberFormatLocal = "\#,##0" '�����u�ʉ�\�̃J���}�ҏW�v

        Else
            '�����ɗ�����A�ǂ��炩�����l�ł͖��������Ƃ������Ƃł�
        End If
    Next r '�A
    
    '== �]�k ==
    '�@�`�C Q005�Ɠ���
    '�D Like ���Z�q�͐��K�\���Ƃ�����ƈႤ���Ǖ֗�
    '   ����̏��Ƃ��Ă͇C�ƇD�ǂ��炪��̕����ǂ��̂��B
    '   ����͖��炩�Ȏd�l���Öق̎d�l�Ƃ������ŁA�D���C�̏��ɂ��܂����B
End Sub
Sub Q007()
    '�\�͈̔͂ƍs���i�[����ꏊ��\��
    Dim rng As Range, r As Range
    '�ŏI�Z���܂ł�\�񂵂��ꏊ�Ɋi�[�iB��Ɍ��o���������Ă���O��j
    Set rng = Range(Range("A2"), Cells.SpecialCells(xlCellTypeLastCell))
    
    '�\��1�s�����o���ď���
    For Each r In rng.Rows
        
        '���t������ꏊ���m�� & ���t�������ꍇ�̖ڈ�쐬
        Dim dt As Date, isTarget As Boolean: isTarget = False
        
        'A�񂪓��t������
        If IsEmpty(r.Cells(1)) Then
            '�󗓂������牽�����Ȃ�
        
        ElseIf IsDate(r.Cells(1)) Then
            '���t�ł���
            dt = r.Cells(1)
            isTarget = True
        
        Else
            '���t��������Ƃ��ē��͂���Ă���\�����l��
            On Error Resume Next
                dt = CDate(r.Cells(1))
                isTarget = (Err.Number = 0)
            On Error GoTo 0
        
        End If
        
        '���t�������疖��������
        If isTarget Then
            r.Cells(2).Value = Format(DateSerial(Year(dt), Month(dt) + 1, 0), "'mmdd")
        End If
        
    Next r
    
End Sub
Sub Q008()
    '�\�͈̔͂ƍs���i�[����ꏊ��\��
    Dim rng As Range, r As Range
    '�ŏI�Z���܂ł�\�񂵂��ꏊ�Ɋi�[�iG��Ɍ��o���������Ă���O��j
    Set rng = Range(Range("B2"), Cells.SpecialCells(xlCellTypeLastCell))
    
    '�\��1�s�����o���ď���
    For Each r In rng.Rows
        
        '���v�_�̓��ꕨ�쐬
        Dim ���v�_ As Integer: ���v�_ = 0
        '50�_�����̖ڈ�쐬
        Dim �ԓ_ As Boolean: �ԓ_ = False
        
        '�s��1��ڂ���5��ڂ����v���Ă���
        Dim i As Long
        For i = 1 To 5
            
            '���v�_�Z�o
            ���v�_ = ���v�_ + r.Cells(i).Value
            
            '50�_��������
            �ԓ_ = �ԓ_ Or r.Cells(i).Value < 50
        
        Next i
        
        '���i��������u���i�v��ݒ�
        If ���v�_ >= 350 And Not �ԓ_ Then
            r.Cells(6) = "���i"
        End If
        
    Next r
    
End Sub
Sub Q009()
    
    '���i�҃V�[�g�A���ѕ\�V�[�g�̏ꏊ�m��
    Dim wsh���i�� As Worksheet, wsh���ѕ\ As Worksheet
    
    '���ѕ\�Ƃ������O�̃V�[�g���A�N�e�B�u�ɂ���
    '�Ȃ�������Excel���牽�炩�̃G���[���o��
    Set wsh���ѕ\ = ActiveWorkbook.Worksheets("���ѕ\")
    
    '�G���[�����̈�W�J
    On Error Resume Next
        
        '�u���i�ҁv�Ƃ������O�̃V�[�g���Ă�ł݂�
        '����������G���[�ŉ������Ȃ����ǃG���[�͖��������
        Set wsh���i�� = ActiveWorkbook.Sheets("���i��")
    
    '�G���[�����̈��
    On Error GoTo 0
    
    '���i�҃V�[�g�������ɌĂ΂ꂽ�ꍇ
    If Not wsh���i�� Is Nothing Then
        '�u�c�������Ă�����e�ؐS
        wsh���i��.Activate
        '�����Ă������m�F���_���Ȃ�I��
        If vbCancel = MsgBox("�����́u���i�ҁv�V�[�g���폜���܂��B", vbOKCancel) Then Exit Sub
        '�_�C�A���O�����ō��̃V�[�g�͏����Ă��܂�
        Application.DisplayAlerts = False
        wsh���i��.Delete
        Application.DisplayAlerts = True
    End If
    
    '�܂�����ȍ��i�҃V�[�g�����
    Set wsh���i�� = ActiveWorkbook.Worksheets.Add
    wsh���i��.Name = "���i��"
    
    '���i�҃��X�g�i�^�u��؂�j
    Dim ���i�ҒB As String
    
    '���ѕ\�V�[�g���̕\�ƍs���i�[����ꏊ��\��
    Dim rng As Range, r As Range
    '���łɍ��i�Ґ����o���Ă����˂�
    Dim i As Long: i = 0
    '���ѕ\�V�[�g���̍ŏI�Z���܂ł�\�񂵂��ꏊ�Ɋi�[
    Set rng = Range(wsh���ѕ\.Range("A2"), wsh���ѕ\.Cells.SpecialCells(xlCellTypeLastCell))
    
    '���ѕ\�V�[�g���̕\����1�s�����o���ď���
    For Each r In rng.Rows
        '���i�������獇�i�҃��X�g�ɒǉ�
        If r.Cells(7).Text = "���i" Then
            ���i�ҒB = ���i�ҒB & r.Cells(1).Text & vbTab
            i = i + 1
        End If
    Next r
    
    '���i�҂Ȃ��̏ꍇ�A�����I��
    If i = 0 Then Exit Sub
    
    '���i�҃��X�g�����i�҃V�[�g�ɒǉ�
    With wsh���i��
        '�z��I�ɉ����тȂ̂ŁA�R�s�y�ŏc���тɕύX
        .Range("A1").Resize(1, i) = Split(���i�ҒB, vbTab)
        .Range("A1").Resize(1, i).Copy
        .Range("A2").PasteSpecial Transpose:=True
        .Range("A1").EntireRow.Delete
    End With
        
    '�����ڂ𐮂���
    wsh���i��.Activate
    wsh���i��.Range("A1").Activate
    
    '== �]�k ==
    '�u���ѕ\�v��������Excel�̃G���[���o�܂�
    'Excel�̃G���[�����̂܂܎g������͎̂g���Ă��܂��i�V�[�g�����\��������Ƃ��j
    'Application.DisplayAlerts ���I�t�ɂ���̂͂�����ƕ|���B�r���ŗ�������ʓ|�B
    '���{��ϐ����A��a�����邯�Ǔ�p����}�V���Ȃ��āc
    '���{��ϐ����ł����Ɂuwsh�v�Ƃ��t���Ă����ƕ⊮(Shift+Tab)�œ��͂ł���̂Ŋy
    
End Sub
Sub Q010()
    
    '�󒍃V�[�g�̏ꏊ�m��
    Dim wsh As Worksheet
    
    '�󒍂Ƃ������O�̃V�[�g���Ă�
    Set wsh = ActiveWorkbook.Worksheets("��")
    
    '�󒍃V�[�g���̕\�ƍs���i�[����ꏊ��\��
    Dim rng As Range, r As Range
    '�󒍃V�[�g���̍ŏI�Z���܂ł�\�񂵂��ꏊ�Ɋi�[
    Set rng = Range(wsh.Range("A2"), wsh.Cells.SpecialCells(xlCellTypeLastCell))
    
    '�󒍃V�[�g���̕\��������1�s�����o���ď���
    Dim i As Long
    For i = rng.Rows.Count To 1 Step -1
        
        '�󒍐����󗓂łȂ��ꍇ
        If Not IsEmpty(rng.Cells(i, 3)) Then
            '�������Ȃ�
        
        '�󒍐����󗓂����l���Ɂu�폜�v�܂��́u�s�v�v�̕������܂܂�Ă���ꍇ
        ElseIf rng.Cells(i, 4).Text Like "*�폜*" _
            Or rng.Cells(i, 4).Text Like "*�s�v*" Then
            '�s�S�̂��폜
            rng.Cells(i, 4).EntireRow.Delete
        
        Else
            '�������Ȃ�
        End If
    
    Next i
        
    '�����ڂ𐮂���
    wsh.Activate
    wsh.Range("A1").Activate
    
    '== �]�k ==
    
    '�E���l���̔���� Select Case �ɂ��悤�������܂������A
    '�@��������� Delete ��2�ӏ��ɔ������邽�߂�߂܂����B
    '�@�Ƃ������������� Like ���g���Ȃ��̂ˁB
    
    '�EForEach �ŏ�̍s���瑀�삷��ƃA�W���p�[�Ȃ̂�
    '�@ForNext �ŉ��̍s���珈������悤�ɂ��܂����B
    '�@��̍s��������Ă����ƁA2�s�A���ō폜�Ώۂ̏ꍇ�ɃA�W���p�[���܂��B
    
End Sub
