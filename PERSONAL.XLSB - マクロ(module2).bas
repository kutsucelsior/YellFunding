Enum ClrIdx

    '�񋓌^�ϐ�
    '�������^�̒l�����ݒ�ł��Ȃ�
    '���W���[�����x���ł�����`�ł��Ȃ��@�ˁ@�錾�Z�N�V�����Œ�`����
    
    �� = 2              '�����������p
    �V�[�O���[�� = 50   '����
    ���C�� = 43         '�\��
    �S�[���h = 44       'S����
    �����I�����W = 45   'N����
    �I�����W = 46       '�J��
    �� = 3              '�m��
    ���F = 8            '����
    G�� = 29            'CXL

End Enum

Sub �z�z�u�b�N�쐬(strOwrName As String, strFltCnd As String)

    Dim wbkPrg As Workbook
    Dim wbkOwr As Workbook
    Dim shtPrg As Worksheet
    Dim shtOwr As Worksheet
    Dim strEndRow As String        '�ŏI�s�̍s�ԍ�
    Dim vntFltCnd As Variant       '�t�B���^�i���ݏ����F�z��
    
    'On Error GoTo Err1
    
    Application.ScreenUpdating = False
    
    '�����i���񍐃u�b�N
    Set wbkPrg = ActiveWorkbook
    Set shtPrg = ActiveSheet
    
    '����MR�ʃu�b�N(���[�N�u�b�N�̐V�K�쐬)
    Set wbkOwr = Workbooks.Add(Template:=xlWBATWorksheet)    'Template:=xlWBATWorksheet �� �V�[�g 1 �����̃u�b�N(�u�b�N���́h"Sheet1")���쐬�ł���
    Set shtOwr = wbkOwr.Worksheets(1)                     'ActiveWorkBook �ƋL�q���Ȃ��Ă������悤�� WorkBook �I�u�W�F�N�g�ɃZ�b�g���Ă���
    
                    '���O��t���ĕۑ����邱�Ƃɂ��A�u�b�N�̖��O��ύX����@�����[�N�u�b�N��Name�v���p�e�B�ł̓u�b�N�̖��O��ύX�o���Ȃ�
    wbkOwr.SaveAs Filename:=wbkPrg.Path & "\" & Left(wbkPrg.Name, Len(wbkPrg.Name) - 4) & "�y" & strOwrName & "�z.xls", _
                     FileFormat:=xlWorkbookNormal   'Excel 97-2003 �u�b�N�`���ŕۑ�
    shtOwr.Name = shtPrg.Name
    
    '�����i���񍐃u�b�N
    shtPrg.Activate
    strEndRow = shtPrg.Cells(Rows.Count, 1).End(xlUp).Row    'A��ڂ̈�ԉ��̃Z�� Cells(Rows.Count, 1)
    vntFltCnd = Split(strFltCnd, ",")   '�J���}��؂�̃t�B���^�i���ݏ����u������v���t�B���^�i���ݏ����u�z��v�ɕϊ�����
    
        'AutoFilter�̈���Criteria1                  �ˁ@�t�B���^�i���ݏ���
        'AutoFilter�̈���Criteria1:=Array(***)      �ˁ@�t�B���^�i���ݏ������R�ȏ�̏ꍇ�́ACriteria1�ɔz����w�肷��
        'AutoFilter�̈���Operator:=xlFilterValues   �ˁ@���̂Ƃ��ACriteria1�ɔz����w�肷�邽�߂ɕt�������Ďw�肷��
    
    Select Case UBound(vntFltCnd) 'UBound(vntFltCnd)�́A�v�f��-1
    Case 0  '���̏ꍇ�A�X���ƃo�O �� �Ȃ����W����Ȃ��ƁA�u�S���ҁv�łȂ��E�ׂ�́u�\�����v�ɂȂ�
        shtPrg.Range("$A$1:$BS$" & strEndRow).AutoFilter Field:=8, Criteria1:=vntFltCnd, Operator:=xlFilterValues
    Case Else
        shtPrg.Range("$A$1:$BS$" & strEndRow).AutoFilter Field:=9, Criteria1:=vntFltCnd, Operator:=xlFilterValues
    End Select
    
    shtPrg.Range("A1").CurrentRegion.Copy
    
    '����MR�ʃu�b�N
    shtOwr.Activate
    shtOwr.Paste                     '�\��t����Range�ł͂Ȃ��AWorksheet�ɑ΂��čs��
    
    Application.CutCopyMode = False     '�R�s�[����͈�(���R�s�[���ɂ���)�������A�_�ł����_���͈̔͂����������
    
    '�R�~�b�V��������폜
    Columns("CB:CB").Delete Shift:=xlToLeft '�z�z�u�b�N�ɂ̓R�~�b�V�������܂߂Ȃ�
    
    shtOwr.Rows("1:1").AutoFilter
    shtOwr.Range("K1").Select
    
    '�����i���񍐃u�b�N
    shtPrg.Activate
    ActiveSheet.AutoFilterMode = False                  '�t�B���^�������� �� ��Őݒ肳�ꂽ����������
    shtPrg.Range("$B$1:$BS$43").AutoFilter Field:=8  'AutoFilter�̈���Criteria1���w�肵�Ȃ� �� �t�B���^���i��Ȃ�(�S�I��)
    shtPrg.Range("K1").Select
    
'    '����MR�ʃu�b�N�t�@�C��
'    shtOwr.Activate
    
    Exit Sub
    
Err1:
    MsgBox "MR�ʃu�b�N�쐬()" & vbCrLf & _
           "�G���[�ԍ�:" & Err.Number & vbCrLf & _
           "�G���[�̎��:" & Err.Description, vbExclamation
           
End Sub

Sub �f�[�^�͈̓Z�b�g(strSheetName As String, strRangeName As String, intTrgRow As Integer, intTrgColumn As Integer)

    'A1����E���[�̃f�[�^�͈͂ɖ��O"Data_Range"���Z�b�g����
    '����intTrgRow    : intTrgRow��ڂ̗�́A�ŉ��[�̃Z���܂Œl�����݂�����I��
    '����intTrgColumn : intTrgColumn�s�ڂ̍s�́A�ŉE�[�̃Z���܂Œl�����݂�����I��(�ʏ�͂P�s�ڂ̃w�b�_�[�s)
    
    Dim strEndRow As Integer        '�ŏI�s�̍s�ԍ�
    Dim strEndColumn As Integer     '�ŏI��̗�ԍ�
    
    'On Error GoTo Err1
    
    Application.ScreenUpdating = False
    
    '��������strSheetName�Ŏw�肳�ꂽ�V�[�g(�S���҂Ȃ�)
    Sheets(strSheetName).Activate
    Sheets(strSheetName).Range("A1").Select     '�f�[�^�͈͊O�̃Z���ɃJ�[�\��������ƁA�����̋��������������Ȃ�@�ˁ@A1�ɃJ�[�\���������Ă����Ζ�薳��
    strEndRow = Cells(Rows.Count, intTrgRow).End(xlUp).Row                      'intTrgRow��ڂ̗�̈�ԉ��̃Z�� Cells(Rows.Count, intTrgRow)
    strEndColumn = Cells(intTrgColumn, Columns.Count).End(xlToLeft).Column      'intTrgColumn�s�ڂ̍s�̈�ԉE�̃Z�� Cells(intTrgColumn, Columns.Count)
    
    'A1����E���[�̃f�[�^�͈͂ɖ��O"Data_Range"���Z�b�g����
    Sheets(strSheetName).Names.Add Name:=strRangeName, RefersToR1C1:="=" & strSheetName & "!R1C1:R" & strEndRow & "C" & strEndColumn
    Sheets(strSheetName).Names(strRangeName).Comment = ""
    
'        '���̃v���V�[�W���P�̂Ŏg���ꍇ�̓R�����g�A�E�g
'        '�J�[�\���ʒu��������
'        Range("A1").Select
'        ActiveWindow.ScrollRow = 1
'        ActiveWindow.ScrollColumn = 1
    
    Exit Sub
    
Err1:
    MsgBox "�f�[�^�͈̓Z�b�g()" & vbCrLf & _
           "�G���[�ԍ�:" & Err.Number & vbCrLf & _
           "�G���[�̎��:" & Err.Description, vbExclamation
           
End Sub

Sub AWK�i���񍐑O����()

    Dim wbkPrg As Workbook
    Dim objCnfWbk As Workbook
    Dim shtPrg As Worksheet
    Dim strEndRow As Integer    '�ŏI�s�̍s�ԍ�
    Dim strIspClm As String     '�񖼂̃A���t�@�x�b�g
    
    'On Error GoTo Err1
    
    'Application.Cursor = xlWait                      '�}�E�X�J�[�\���������v�ɂɂ���
    Application.ScreenUpdating = False
    
    '�ȍ~�̏����ŃA�N�e�B�u�u�b�N���ς��O�ɁA
    '�v���V�[�W���ďo�����̃A�N�e�B�u�u�b�N�ƃA�N�e�B�u�V�[�g���擾���Ă���
    Set wbkPrg = ActiveWorkbook
    Set shtPrg = ActiveSheet
    
    '���ёւ��@ISP�_��ԍ�(����)
    'Order �� xlAscending(����),xlDescending(�~��)
    strEndRow = Cells(Rows.Count, 1).End(xlUp).Row  'A��ڂ̈�ԉ��̃Z�� Cells(Rows.Count, 1)
    shtPrg.Sort.SortFields.Clear
    
    'ISP�_��ԍ��̗񖼃A���t�@�x�b�g���擾����
    strIspClm = Cells.Find(What:="ISP�_��ԍ�", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, MatchByte:=False, SearchFormat:=False).Address
    strIspClm = Mid(strIspClm, 2, 1)
    
    shtPrg.Sort.SortFields.Add Key:=Range(strIspClm & "2:" & strIspClm & strEndRow) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With shtPrg.Sort
        .SetRange Range("A1:BO" & strEndRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    '�O�񏈗��ŗ񂪐�������Ă���΍폜
    If shtPrg.Range("A1").FormulaR1C1 = "�W�v��" _
        And shtPrg.Range("I1").FormulaR1C1 = "�S����" _
        And shtPrg.Range("J1").FormulaR1C1 = "�\����" _
        And shtPrg.Range("K1").FormulaR1C1 = "�t�F�[�Y" _
        And shtPrg.Range("CB1").FormulaR1C1 = "�R�~�b�V����" Then
        Range("A:A,I:I,J:J,K:K,CB:CB").Delete Shift:=xlToLeft
    End If
    
    '�ǉ��t�B�[���h��}��
    shtPrg.Select
    shtPrg.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    shtPrg.Range("A1").FormulaR1C1 = "�W�v��"
    shtPrg.Columns("I:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    shtPrg.Range("I1").FormulaR1C1 = "�S����"
    shtPrg.Columns("J:J").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    shtPrg.Range("J1").FormulaR1C1 = "�\����"
    shtPrg.Columns("K:K").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    shtPrg.Range("K1").FormulaR1C1 = "�t�F�[�Y"
    shtPrg.Columns("CB:CB").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    shtPrg.Columns("CB:CB").NumberFormatLocal = "\#,##0;\-#,##0"
    shtPrg.Range("CB1").FormulaR1C1 = "�R�~�b�V����"
    
    '�u�b�N �J��
    '�u��荞�݌��v
    Set objCnfWbk = Workbooks.Open(wbkPrg.Path & "\NURO�i���񍐐ݒ�.xls", Password:="0911")
    
    '�V�[�g ��荞��
    '�u�S���ҁv
    If ExistSheet("�S����", wbkPrg) Then '���Ɏ�荞��ł���΍폜
        Application.DisplayAlerts = False
        wbkPrg.Worksheets("�S����").Delete
        Application.DisplayAlerts = True
    End If
    objCnfWbk.Worksheets("�S����").Copy After:=wbkPrg.Worksheets(wbkPrg.Worksheets.Count)
    '�u�L�����y�[���v
    If ExistSheet("�L�����y�[��", wbkPrg) Then   '���Ɏ�荞��ł���΍폜
        Application.DisplayAlerts = False
        wbkPrg.Worksheets("�L�����y�[��").Delete
        Application.DisplayAlerts = True
    End If
    objCnfWbk.Worksheets("�L�����y�[��").Copy After:=wbkPrg.Worksheets(wbkPrg.Worksheets.Count)
    '�u�z�z�v
    If ExistSheet("�z�z", wbkPrg) Then '���Ɏ�荞��ł���΍폜
        Application.DisplayAlerts = False
        wbkPrg.Worksheets("�z�z").Delete
        Application.DisplayAlerts = True
    End If
    objCnfWbk.Worksheets("�z�z").Copy After:=wbkPrg.Worksheets(wbkPrg.Worksheets.Count)
    
    '�ݒ�t�@�C����ۑ����Ȃ��ŕ���
    objCnfWbk.Close saveChanges:=False
    
    '�v���V�[�W���ďo�����̃A�N�e�B�u�u�b�N�ɖ߂�
    shtPrg.Activate
    shtPrg.Range("K1").Select
    
    Exit Sub
    
Err1:
    MsgBox "AWK�i���񍐑O����()" & vbCrLf & _
           "�G���[�ԍ�:" & Err.Number & vbCrLf & _
           "�G���[�̎��:" & Err.Description, vbExclamation
           
End Sub

Function SQL����() As String

    Dim shtPrg As Worksheet
    
    'On Error GoTo Err1
    
    Application.ScreenUpdating = False
    
    '�ȍ~�̏����ŃA�N�e�B�u�u�b�N���ς��O�ɁA
    '�v���V�[�W���ďo�����̃A�N�e�B�u�V�[�g���擾���Ă���
    Set shtPrg = ActiveWorkbook.ActiveSheet
    
    '���\�߁A�f�[�^�͈͂ɖ��O��t���Ă���
    '���u�b�N��̃f�[�^�͈͂ł͂Ȃ��A�V�[�g��̃f�[�^�͈͂ɖ��O��t����
    Call �f�[�^�͈̓Z�b�g("SP_WORK", "Data_Range", 8, 1)   '�󔒂�������ƍs�́A�ԍ����Z�b�g�@�ˁ@�W��ځFISP�_��ԍ��A�P�s�ځF�񖼃w�b�_�[�s
    Call �f�[�^�͈̓Z�b�g("�S����", "Owner_Range", 1, 1)   '�󔒂�������ƍs�́A�ԍ����Z�b�g�@�ˁ@�P��ځFISP�_��ԍ��A�P�s�ځF�񖼃w�b�_�[�s
    Call �f�[�^�͈̓Z�b�g("�L�����y�[��", "Campaign_Range", 1, 1)   '�󔒂�������ƍs�́A�ԍ����Z�b�g�@�ˁ@�P��ځFISP�_��ԍ��A�P�s�ځF�񖼃w�b�_�[�s
    
    SQL���� = "select " & _
                    "[SP_WORK$Data_Range].[�\����], " & _
                    "[�S����$Owner_Range].[�S����], " & _
                    "[SP_WORK$Data_Range].[�������], " & _
                    "[SP_WORK$Data_Range].[So-net�H���\���], " & _
                    "[SP_WORK$Data_Range].[NTT�H���\���], " & _
                    "[SP_WORK$Data_Range].[So-net�H����], " & _
                    "[SP_WORK$Data_Range].[NTT�H����], " & _
                    "[SP_WORK$Data_Range].[NURO������J�ʏ�����], " & _
                    "[SP_WORK$Data_Range].[���Ϗ��m���], " & _
                    "[SP_WORK$Data_Range].[�L�����Z����], " & _
                    "[SP_WORK$Data_Range].[���A����d�b�ԍ�], " & _
                    "[�L�����y�[��$Campaign_Range].[�R�~�b�V����] " & _
              "from  [SP_WORK$Data_Range], " & _
                    "[�S����$Owner_Range], " & _
                    "[�L�����y�[��$Campaign_Range] " & _
              "Where [SP_WORK$Data_Range].[ISP�_��ԍ�] = [�S����$Owner_Range].[ISP�_��ԍ�] " & _
              "And   [SP_WORK$Data_Range].[�㗝�X�R�[�h] = [�L�����y�[��$Campaign_Range].[�L�����y�[���R�[�h] " & _
              "Order by [SP_WORK$Data_Range].[ISP�_��ԍ�] "
    
    '�v���V�[�W���ďo�����̃A�N�e�B�u�u�b�N�ɖ߂�
    shtPrg.Activate
    shtPrg.Range("K1").Select
    
    Exit Function
    
Err1:
    MsgBox "SQL����()" & vbCrLf & _
           "�G���[�ԍ�:" & Err.Number & vbCrLf & _
           "�G���[�̎��:" & Err.Description, vbExclamation
           
End Function

Sub AWK�i����()

    Dim objCn As New ADODB.Connection
    Dim objRs  As ADODB.Recordset
    Dim strSQL As String
    Dim strBuf As String
    Dim intRow As Integer
    Dim dtm�W�v�� As Date
    Dim str�t�F�[�Y As String
    Dim wbkPrg As Workbook
    Dim shtPrg As Worksheet
    
    'On Error GoTo Err1
    
    '�ȍ~�̏����ŃA�N�e�B�u�u�b�N���ς��O�ɁA
    '�v���V�[�W���ďo�����̃A�N�e�B�u�u�b�N�ƃA�N�e�B�u�V�[�g���擾���Ă���
    Set wbkPrg = ActiveWorkbook
    Set shtPrg = ActiveSheet
    
    Call AWK�i���񍐑O����
    
    With objCn
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & wbkPrg.Path & "\" & wbkPrg.Name & ";" & _
         "Extended Properties=Excel 8.0;"
        .Open
    End With
    
    'SQL����() �� �f�[�^�͈͐ݒ�������ɍs��
    strSQL = SQL����
    Set objRs = objCn.Execute(strSQL)
    
'                            Debug.Print "���R�[�h��" & objRs.Fields.Count
'
'                            '��ԗ񖼃��X�g���_���v(�C�~�f�B�G�C�g�E�B���h�E��)
'                            Dim i
'                            i = 0
'                            For i = 0 To objRs.Fields.Count - 1
'                                Debug.Print i & " " & objRs.Fields(i).Name
'                            Next
'
'                            '�񖼃w�b�_�[���_���v(�C�~�f�B�G�C�g�E�B���h�E��)
'                            strBuf = objRs.Fields(0).Name _
'                             & "," & objRs.Fields(1).Name _
'                             & ",�\����" _
'                             & ",�t�F�[�Y" _
'                             & "," & objRs.Fields(2).Name _
'                             & "," & objRs.Fields(3).Name _
'                             & "," & objRs.Fields(4).Name _
'                             & "," & objRs.Fields(5).Name _
'                             & "," & objRs.Fields(6).Name _
'                             & "," & objRs.Fields(7).Name _
'                             & "," & objRs.Fields(8).Name _
'                             & "," & objRs.Fields(9).Name _
'                             & "," & IIf(strMode = "admin", objRs.Fields(10), "")
'                             Debug.Print strBuf
    
    intRow = 1   '�s��
    dtm�W�v�� = CDate(Format(Mid(wbkPrg.Name, 9, 8), "@@@@/@@/@@"))
    Do While objRs.EOF = False
    
        '�t�F�[�Y���擾
        str�t�F�[�Y = get�t�F�[�Y( _
                                    CStr(dtm�W�v��), _
                                    IIf(IsNull(objRs![�\����]), "", objRs![�\����]), _
                                    IIf(IsNull(objRs![So-net�H���\���]), "", objRs![So-net�H���\���]), _
                                    IIf(IsNull(objRs![NTT�H���\���]), "", objRs![NTT�H���\���]), _
                                    IIf(IsNull(objRs![So-net�H����]), "", objRs![So-net�H����]), _
                                    IIf(IsNull(objRs![NTT�H����]), "", objRs![NTT�H����]), _
                                    IIf(IsNull(objRs![NURO������J�ʏ�����]), "", objRs![NURO������J�ʏ�����]), _
                                    IIf(IsNull(objRs![���Ϗ��m���]), "", objRs![���Ϗ��m���]), _
                                    IIf(IsNull(objRs![�L�����Z����]), "", objRs![�L�����Z����]) _
                                )
        '�\�������擾
        str�\���� = Chk��(objRs![�\����], CStr(dtm�W�v��))
        
'                            '�f�[�^���_���v(�C�~�f�B�G�C�g�E�B���h�E)
'                            strBuf = objRs![�\����] _
'                            & "," & objRs![�S����] _
'                            & "," & str�\���� & "�\��" _
'                            & "," & str�t�F�[�Y _
'                            & "," & objRs![�������] _
'                            & "," & objRs![So-net�H���\���] _
'                            & "," & objRs![NTT�H���\���] _
'                            & "," & objRs![So-net�H����] _
'                            & "," & objRs![NTT�H����] _
'                            & "," & objRs![NURO������J�ʏ�����] _
'                            & "," & objRs![���Ϗ��m���] _
'                            & "," & objRs![�L�����Z����] _
'                            & "," & IIf(strMode = "admin", objRs![�R�~�b�V����], "")
'                            Debug.Print strBuf
        
        '�V�[�g�֏�������
        shtPrg.Range("A" & intRow + 1) = CStr(dtm�W�v��)
        shtPrg.Range("I" & intRow + 1) = objRs![�S����]
        shtPrg.Range("J" & intRow + 1) = str�\���� & "�\��"
        shtPrg.Range("K" & intRow + 1) = str�t�F�[�Y
        If Not IsNull(objRs![���A����d�b�ԍ�]) Then
            If InStr(objRs![���A����d�b�ԍ�], "-") = 0 Then
                shtPrg.Range("N" & intRow + 1) = "0" & Mid(objRs![���A����d�b�ԍ�], 1, 1) & "-" & Mid(objRs![���A����d�b�ԍ�], 2, 4) & "-" & Mid(objRs![���A����d�b�ԍ�], 6, 4)
            End If
        End If
        shtPrg.Range("CB" & intRow + 1) = objRs![�R�~�b�V����]
        
        intRow = intRow + 1
        objRs.MoveNext  '���̃��R�[�h�ֈړ�
    Loop
    
    objRs.Close
    Set objRs = Nothing
    
    
    Call AWK�i���񍐌㏈��
    
    '�z�z����
    '������������������������������������������������
    
    Call �f�[�^�͈̓Z�b�g("�z�z", "Dist_Range", 1, 1)
    
    '�v���V�[�W���ďo�����̃A�N�e�B�u�u�b�N�ɖ߂�
    shtPrg.Activate
    shtPrg.Range("K1").Select
    
    strSQL = "select " & _
                    "[�z�z$Dist_Range].[�z�z��], " & _
                    "[�z�z$Dist_Range].[�z�z����] " & _
              "from  [�z�z$Dist_Range]"
    
    Set objRs = objCn.Execute(strSQL)
    Do While objRs.EOF = False
    
        Call �z�z�u�b�N�쐬(objRs![�z�z��], objRs![�z�z����])
        
        objRs.MoveNext  '���̃��R�[�h�ֈړ�
    Loop
    
    If ExistSheet("�z�z") Then
        Application.DisplayAlerts = False
        Worksheets("�z�z").Delete
        Application.DisplayAlerts = True
    End If
    
    '������������������������������������������������
    
    objRs.Close
    objCn.Close
    
    Set objRs = Nothing
    Set objCn = Nothing
    
    Exit Sub
    
Err1:
    MsgBox "AWK�i����()" & vbCrLf & _
           "�G���[�ԍ�:" & Err.Number & vbCrLf & _
           "�G���[�̎��:" & Err.Description, vbExclamation
           
End Sub

Sub AWK�i���񍐌㏈��()

    Dim shtPrg As Worksheet
    
    'On Error GoTo Err1
    
    Application.ScreenUpdating = False
    
    Set shtPrg = ActiveSheet
    
    Call �t�F�[�Y_�����t����_�F�ݒ�
    
    '�V�[�g�����݂��Ă���΍폜
    If ExistSheet("�S����") Then
        Application.DisplayAlerts = False
        Worksheets("�S����").Delete
        Application.DisplayAlerts = True
    End If
    If ExistSheet("�L�����y�[��") Then
        Application.DisplayAlerts = False
        Worksheets("�L�����y�[��").Delete
        Application.DisplayAlerts = True
    End If
    
    '"SP_WORK"��I��
    shtPrg.Select
    
    '�I�[�g�t�B���^�ݒ�
    If ActiveSheet.AutoFilterMode = False Then
        shtPrg.Rows("1:1").AutoFilter
    End If
    
    '�񕝂��œK��
    shtPrg.Columns("A:A").EntireColumn.AutoFit   '�W�v��(�ǉ�������)
    shtPrg.Columns("H:H").EntireColumn.AutoFit   'ISP�_��ԍ�
    shtPrg.Columns("I:K").EntireColumn.AutoFit   '�S����,�\����,�t�F�[�Y�ǉ�������
    shtPrg.Columns("N:N").EntireColumn.AutoFit   '���A����d�b�ԍ�
    shtPrg.Columns("S:S").EntireColumn.AutoFit   '�\����
    shtPrg.Columns("Y:AD").EntireColumn.AutoFit   'So-net�H���\���,NTT�H���\���,So-net�H����,NTT�H����,NURO������J�ʏ�����,�L�����Z����
    shtPrg.Columns("AF:AF").EntireColumn.AutoFit   '�\����
    
    '�����I����J�[�\���ʒu
    shtPrg.Range("K1").Select
    
    Application.Cursor = xlDefault                      '�}�E�X�J�[�\���������v���猳�ɂɂ���
    
    Exit Sub
    
Err1:
    MsgBox "AWK�i���񍐌㏈��()" & vbCrLf & _
           "�G���[�ԍ�:" & Err.Number & vbCrLf & _
           "�G���[�̎��:" & Err.Description, vbExclamation
           
End Sub

Sub �t�F�[�Y_�����t����_�F�ݒ�()

    'On Error GoTo Err1
    
    Sheets("SP_WORK").Activate
    
    Columns("K:K").Select
    
                '���������@�����t������
                'FormatConditions.Add�@�����t��������ǉ�
                'Type:=xlTextString �� ����̕�����
                'TextOperator:=xlContains �� ���̒l���܂� �Z���̒l���̒l�ɓ�����
                'String:=�@���Ŏw��
                
                'Type:=xlCellValue �� �Z���̒l
                'Operator:=xlEqual �� ���̒l�ɓ�����
                'Formula1:=�@���Ŏw��
                
                '�D�揇�ʂ��P�ʂɂ��ď����t��������ǉ�����
                'Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                
                '�����t�������́w�����𖞂����ꍇ�͒�~�x�@�ˁ@��������u���́v(�������ɐݒ肵��)�����𒲂ׂ邩�ǂ�����ݒ�
                'Selection.FormatConditions(1).StopIfTrue = True
    
    '�\��
    Selection.FormatConditions.Add Type:=xlTextString, String:="����", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.�V�[�O���[��  '50
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    '�\��
    Selection.FormatConditions.Add Type:=xlTextString, String:="�\��", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.���C��  '43
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    'S����
    Selection.FormatConditions.Add Type:=xlTextString, String:="S����", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.�S�[���h '44
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    'N����
    Selection.FormatConditions.Add Type:=xlTextString, String:="N����", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.�����I�����W '45
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    '�J��
    Selection.FormatConditions.Add Type:=xlTextString, String:="�J��", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.�I�����W '46
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    '�m��
    Selection.FormatConditions.Add Type:=xlTextString, String:="�m��", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ColorIndex = ClrIdx.�� '2
    End With
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.�� '3
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    '����
    Selection.FormatConditions.Add Type:=xlTextString, String:="����", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.���F '8
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    'CXL
    Selection.FormatConditions.Add Type:=xlTextString, String:="CXL", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ColorIndex = ClrIdx.�� '2
    End With
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.G�� '29
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    Exit Sub
    
Err1:
    MsgBox "�t�F�[�Y_�����t����_�F�ݒ�()" & vbCrLf & _
           "�G���[�ԍ�:" & Err.Number & vbCrLf & _
           "�G���[�̎��:" & Err.Description, vbExclamation
           
End Sub

Function get�t�F�[�Y(str�W�v�� As String, str�\���� As String, strS�\��� As String, strN�\��� As String, _
                        strS�H���� As String, strN�H���� As String, str�J�ʓ� As String, str���ϊm��� As String, strCXL�� As String) As String
    
    Dim flg_�\���� As String
    Dim flg_SON�\��� As String
    Dim flg_N�\��� As String
    Dim flg_SON�H���� As String
    Dim flg_N�H���� As String
    Dim flg_�J�ʓ� As String
    Dim flg_���ϊm��� As String
    Dim flg_CXL�� As String
    
    'On Error GoTo Err1
    
    Application.ScreenUpdating = False
    
    flg_�\���� = Chk��(str�\����, str�W�v��)
    flg_SON�\��� = Chk��(strS�\���, str�W�v��)
    flg_N�\��� = Chk��(strN�\���, str�W�v��)
    flg_SON�H���� = Chk��(strS�H����, str�W�v��)
    flg_N�H���� = Chk��(strN�H����, str�W�v��)
    flg_�J�ʓ� = Chk��(str�J�ʓ�, str�W�v��)
    flg_���ϊm��� = Chk��(str���ϊm���, str�W�v��)
    flg_CXL�� = Chk��(strCXL��, str�W�v��)
    
    Select Case flg_�\����                                              '�����t�F�[�Y�̏W�v�ΏہF�\�������O�X���ȍ~
    Case "����", "�O��", "�O�X��"                                       '����CXL�̏W�v�ΏہF�\�������O�X���ȍ~
        Select Case flg_CXL��
        Case "����"
            get�t�F�[�Y = "CXL" '�����́A����CXL
        Case "��"                                                     '�����J�ʂ܂��͊m��̏W�v�ΏہFCXL������
            Select Case flg_�J�ʓ�
            Case "����"
                Select Case flg_���ϊm���
                Case "��"
                    get�t�F�[�Y = "�����J��"
                Case "�O�X�X���ȑO", "�O�X��", "�O��", "����", "����", "�ė����ȍ~"
                    get�t�F�[�Y = "�����m��"
                Case Else
                    get�t�F�[�Y = ""
                End Select
                
            Case "�O��"
                Select Case flg_���ϊm���
                Case "��"
                    get�t�F�[�Y = "�O���J��"
                Case "�O�X�X���ȑO", "�O�X��", "�O��", "����", "����", "�ė����ȍ~"
                    'get�t�F�[�Y = "�O���m��"
                    get�t�F�[�Y = ""               '�W�v�ΏۊO�F�O���m��
                Case Else
                    get�t�F�[�Y = ""
                End Select
                
            Case "�O�X��"
                Select Case flg_���ϊm���
                Case "��"
                    get�t�F�[�Y = "�O�X���J��"
                Case "�O�X�X���ȑO", "�O�X��", "�O��", "����", "����", "�ė����ȍ~"
                    'get�t�F�[�Y = "�O�X���m��"
                    get�t�F�[�Y = ""               '�W�v�ΏۊO�F�O�X���m��
                Case Else
                    get�t�F�[�Y = ""
                End Select
                
            Case "��"                                             '����N�����̏W�v�ΏہF�J�ʂ���
                Select Case flg_N�H����
                Case "����"
                    get�t�F�[�Y = "����N����"
                Case "�O��"
                    get�t�F�[�Y = "�O��N����"
                Case "�O�X��"
                    get�t�F�[�Y = "�O�X��N����"
                Case "��"                                         '����S�����̏W�v�ΏہF�J�ʂ���
                    Select Case flg_SON�H����
                    Case "����"
                        get�t�F�[�Y = "����S����"
                    Case "����"
                        get�t�F�[�Y = "�O��S����"
                    Case "����"
                        get�t�F�[�Y = "�O�X��S����"
                    Case "��"                                     '����SN�\��,����̏W�v�ΏہF�J�ʂ���
                        Select Case flg_N�\���
                        Case "����"
                            Select Case flg_SON�\���
                            Case "����"
                                get�t�F�[�Y = "����SN�\��"          '����S����N�\��
                            Case "����"
                                get�t�F�[�Y = "����SN�\��"          '����S����N�\��
                            Case "�O��"
                                get�t�F�[�Y = "����SN�\��"          '�O��S����N�\��
                            Case "�O�X��"
                                get�t�F�[�Y = "����SN�\��"          '�O�X��S����N�\��
                            Case "��"
                                get�t�F�[�Y = "����N�̂ݗ\��"       '����S����N�\��
                            End Select
                            
                        Case "����"
                            Select Case flg_SON�\���
                            Case "����"
                                get�t�F�[�Y = "����SN�\��"          '����S����N�\��
                            Case "����"
                                get�t�F�[�Y = "����SN�\��"          '����S����N�\��
                            Case "�O��"
                                get�t�F�[�Y = "����SN�\��"          '�O��S����N�\��
                            Case "�O�X��"
                                get�t�F�[�Y = "����SN�\��"          '�O�X��S����N�\��
                            Case "��"
                                get�t�F�[�Y = "����N�̂ݗ\��"       '����S����N�\��
                            End Select
                            
                        Case "�O��"
                            Select Case flg_SON�\���
                            Case "����"
                                get�t�F�[�Y = "����SN�\��"          '����S�O��N�\��
                            Case "����"
                                get�t�F�[�Y = "����SN�\��"          '����S�O��N�\��
                            Case "�O��"
                                get�t�F�[�Y = "�O��SN�\��"          '�O��SN�\��
                            Case "�O�X��"
                                get�t�F�[�Y = "�O��SN�\��"          '�O�X��S�O��N�\��
                            Case "��"
                                get�t�F�[�Y = "�O��N�̂ݗ\��"       '����S�O��N�\��
                            End Select
                            
                        Case "�O�X��"
                            Select Case flg_SON�\���
                            Case "����"
                                get�t�F�[�Y = "����SN�\��"          '����S�O�X��N�\��
                            Case "����"
                                get�t�F�[�Y = "����SN�\��"          '����S�O�X��N�\��
                            Case "�O��"
                                get�t�F�[�Y = "�O��SN�\��"          '�O��S�O�X��N�\��
                            Case "�O�X��"
                                get�t�F�[�Y = "�O�X��SN�\��"        '�O�X��S�O�X��N�\��
                            Case "��"
                                get�t�F�[�Y = "�O�X��N�̂ݗ\��"     '����S�O�X��N�\��
                            End Select
                            
                        Case "��"
                            Select Case flg_SON�\���
                            Case "����"
                                get�t�F�[�Y = "����S�̂ݗ\��"       '����S����N�\��
                            Case "����"
                                get�t�F�[�Y = "����S�̂ݗ\��"       '����S����N�\��
                            Case "�O��"
                                get�t�F�[�Y = "�O��S�̂ݗ\��"       '�O��S����N�\��
                            Case "�O��"
                                get�t�F�[�Y = "�O�X��S�̂ݗ\��"         '�O�X��S����N�\��
                            Case "��"
                                get�t�F�[�Y = "����"                '����S����N�\��
                            End Select
                            
                        End Select
                    Case Else
                        get�t�F�[�Y = ""    '�W�v�ΏۊO�FS�H�������A�s���l
                    End Select
                Case Else
                    get�t�F�[�Y = ""    '�W�v�ΏۊO�FN�H�������A�s���l
                End Select
            Case Else
                get�t�F�[�Y = ""    '�W�v�ΏۊO�F�J�ʓ����A�s���l
            End Select
        Case Else
            get�t�F�[�Y = ""    '�W�v�ΏۊO�FCXL�����A�O���ȑO�܂��͕s���l
        End Select
    Case Else
        get�t�F�[�Y = ""    '�W�v�ΏۊO�F�\�������A�O�X���ȑO�܂��͕s���l
    End Select
    
    Exit Function
    
Err1:
    MsgBox "get�t�F�[�Y()" & vbCrLf & _
           "�G���[�ԍ�:" & Err.Number & vbCrLf & _
           "�G���[�̎��:" & Err.Description, vbExclamation
           
End Function

Function Chk��(strTrgDate As String, strAggDate As String) As String

    'On Error GoTo Err1
    
    Application.ScreenUpdating = False
    
    'strTrgDate : �]�����t
    'strAggDate : �W�v���t
    
    If strTrgDate = "" Then '�]�����t���A�󔒂Ȃ�B�B�B
    
        Chk�� = "��"
        
    '�]�����t���A�󔒂łȂ��Ȃ�B�B�B
    Else
        
        '(�Q�l)�@�����擾 DateSerial(Year("yyyy/mm/dd"), Month("yyyy/mm/dd"), 1)
        '�@�@�@  �����擾 DateSerial(Year("yyyy/mm/dd"), Month("yyyy/mm/dd"), 0)
        
        Select Case CDate(strTrgDate)   '�]�����t���A�ȉ��͈͓̔��Ȃ�B�B�B
        
            '�ė������� �ȍ~
            Case Is >= DateSerial(Year(strAggDate), Month(strAggDate) + 2, 1)
                Chk�� = "�ė����ȍ~"
                
            '�������� ���� �������� �̊�
            Case DateSerial(Year(strAggDate), Month(strAggDate) + 1, 1) To DateSerial(Year(strAggDate), Month(strAggDate) + 2, 0)
                Chk�� = "����"
                
            '�������� ���� �������� �̊�
            Case DateSerial(Year(strAggDate), Month(strAggDate) + 0, 1) To DateSerial(Year(strAggDate), Month(strAggDate) + 1, 0)
                Chk�� = "����"
                
            '�O������ ���� �O������ �̊�
            Case DateSerial(Year(strAggDate), Month(strAggDate) - 1, 1) To DateSerial(Year(strAggDate), Month(strAggDate) + 0, 0)
                Chk�� = "�O��"
                
            '�O�X������ ���� �O�X������ �̊�
            Case DateSerial(Year(strAggDate), Month(strAggDate) - 2, 1) To DateSerial(Year(strAggDate), Month(strAggDate) - 1, 0)
                Chk�� = "�O�X��"
                
            '�O�X������ ���ߋ�
            Case Is < DateSerial(Year(strAggDate), Month(strAggDate) - 2, 1)
                Chk�� = "�O�X�X���ȑO"
                
        End Select
        
    End If
    
    Exit Function
    
Err1:
    MsgBox "Chk��()" & vbCrLf & _
           "�G���[�ԍ�:" & Err.Number & vbCrLf & _
           "�G���[�̎��:" & Err.Description, vbExclamation
           
End Function

Sub ��{�`_Excel��ODBC�ڑ�()

    Dim objCn As New ADODB.Connection
    Dim objRs  As ADODB.Recordset
    Dim strSQL As String
    Dim strBuf As String
    
    'On Error GoTo Err1
    
    With objCn
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & ActiveWorkbook.Path & "\" & ActiveWorkbook.Name & ";" & _
         "Extended Properties=Excel 8.0;"
        .Open
    End With
    
    '�f�[�^�͈͂ɗ\�ߖ��O��t���Ă���(�u�b�N�ł͂Ȃ��A�V�[�g��͈̔͂ɂ���)
    'SQL�̃e�[�u�����́A[�V�[�g��$�f�[�^�͈͖�]�̃t�H�[�}�b�g�Ŏw�肷��
    
    'strSQL = "select * from [" & ActiveSheet.Name & "$Data_Range]"
    strSQL = "select * from [" & ActiveSheet.Name & "$SP_WORK]"
    Set objRs = objCn.Execute(strSQL)
    
        '    '��ԗ񖼃_���v
        '    Dim i
        '    i = 0
        '    For i = 0 To Cells(1, Columns.Count).End(xlToLeft).Column - 2
        '        Debug.Print i & " " & objRs.Fields(i).Name
        '    Next
    
    '�񖼃_���v
    strBuf = objRs.Fields(10).Name & "," & objRs.Fields(23).Name & "," & objRs.Fields(24).Name & "," & objRs.Fields(25).Name & "," & objRs.Fields(26).Name & "," & objRs.Fields(27).Name & "," & objRs.Fields(28).Name
    Debug.Print strBuf
    'ActiveSheet.Range("A40:I40") = Split(strBuf, ",")
    
    '�f�[�^�_���v
    Do While objRs.EOF = False
        Debug.Print objRs!������� & ", " & objRs![S�H���\���] & ", " & objRs![N�H���\���] & ", " & objRs![S�H����] & ", " & objRs![N�H����] & ", " & objRs![NURO������J�ʏ�����] & ", " & objRs![�L�����Z����]
        objRs.MoveNext  '���̃��R�[�h�ֈړ�
    Loop
    
    objRs.Close
    objCn.Close
    
    Set objRs = Nothing
    Set objCn = Nothing
    
    Exit Sub
    
Err1:
    MsgBox "��{�`_Excel��ODBC�ڑ�()" & vbCrLf & _
           "�G���[�ԍ�:" & Err.Number & vbCrLf & _
           "�G���[�̎��:" & Err.Description, vbExclamation
           
End Sub

'���[�N�V�[�g�����݂��邩���ׂ�֐�
Function ExistSheet(strSheetName As String, Optional objWbk As Variant) As Boolean

    Dim objSheet As Object
    
    'On Error GoTo Err1
      
    ExistSheet = False
    'IsMissing�̈�����Variant�^�ł���K�v����
    If IsMissing(objWbk) Then
        For Each objSheet In ActiveWorkbook.Sheets
            If objSheet.Name = strSheetName Then
                ExistSheet = True
                Exit For
            End If
        Next objSheet
    Else
        For Each objSheet In objWbk.Sheets
            If objSheet.Name = strSheetName Then
                ExistSheet = True
                Exit For
            End If
        Next objSheet
    End If
    
    Exit Function
    
Err1:
    MsgBox "ExistSheet()" & vbCrLf & _
           "�G���[�ԍ�:" & Err.Number & vbCrLf & _
           "�G���[�̎��:" & Err.Description, vbExclamation
           
End Function



