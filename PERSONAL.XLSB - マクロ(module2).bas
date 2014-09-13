Enum ClrIdx

    '列挙型変数
    '長整数型の値しか設定できない
    'モジュールレベルでしか定義できない　⇒　宣言セクションで定義する
    
    白 = 2              '白抜き文字用
    シーグリーン = 50   '未定
    ライム = 43         '予定
    ゴールド = 44       'S完了
    薄いオレンジ = 45   'N完了
    オレンジ = 46       '開通
    赤 = 3              '確定
    水色 = 8            '来月
    G紫 = 29            'CXL

End Enum

Sub 配布ブック作成(strOwrName As String, strFltCnd As String)

    Dim wbkPrg As Workbook
    Dim wbkOwr As Workbook
    Dim shtPrg As Worksheet
    Dim shtOwr As Worksheet
    Dim strEndRow As String        '最終行の行番号
    Dim vntFltCnd As Variant       'フィルタ絞込み条件：配列
    
    'On Error GoTo Err1
    
    Application.ScreenUpdating = False
    
    '↓↓進捗報告ブック
    Set wbkPrg = ActiveWorkbook
    Set shtPrg = ActiveSheet
    
    '↓↓MR別ブック(ワークブックの新規作成)
    Set wbkOwr = Workbooks.Add(Template:=xlWBATWorksheet)    'Template:=xlWBATWorksheet ⇒ シート 1 つだけのブック(ブック名は”"Sheet1")が作成できる
    Set shtOwr = wbkOwr.Worksheets(1)                     'ActiveWorkBook と記述しなくてもいいように WorkBook オブジェクトにセットしておく
    
                    '名前を付けて保存することにより、ブックの名前を変更する　※ワークブックのNameプロパティではブックの名前を変更出来ない
    wbkOwr.SaveAs Filename:=wbkPrg.Path & "\" & Left(wbkPrg.Name, Len(wbkPrg.Name) - 4) & "【" & strOwrName & "】.xls", _
                     FileFormat:=xlWorkbookNormal   'Excel 97-2003 ブック形式で保存
    shtOwr.Name = shtPrg.Name
    
    '↓↓進捗報告ブック
    shtPrg.Activate
    strEndRow = shtPrg.Cells(Rows.Count, 1).End(xlUp).Row    'A列目の一番下のセル Cells(Rows.Count, 1)
    vntFltCnd = Split(strFltCnd, ",")   'カンマ区切りのフィルタ絞込み条件「文字列」をフィルタ絞込み条件「配列」に変換する
    
        'AutoFilterの引数Criteria1                  ⇒　フィルタ絞込み条件
        'AutoFilterの引数Criteria1:=Array(***)      ⇒　フィルタ絞込み条件が３つ以上の場合は、Criteria1に配列を指定する
        'AutoFilterの引数Operator:=xlFilterValues   ⇒　このとき、Criteria1に配列を指定するために付け加えて指定する
    
    Select Case UBound(vntFltCnd) 'UBound(vntFltCnd)は、要素数-1
    Case 0  'この場合、９だとバグ ⇒ なぜか８じゃないと、「担当者」でなく右隣りの「申込月」になる
        shtPrg.Range("$A$1:$BS$" & strEndRow).AutoFilter Field:=8, Criteria1:=vntFltCnd, Operator:=xlFilterValues
    Case Else
        shtPrg.Range("$A$1:$BS$" & strEndRow).AutoFilter Field:=9, Criteria1:=vntFltCnd, Operator:=xlFilterValues
    End Select
    
    shtPrg.Range("A1").CurrentRegion.Copy
    
    '↓↓MR別ブック
    shtOwr.Activate
    shtOwr.Paste                     '貼り付けはRangeではなく、Worksheetに対して行う
    
    Application.CutCopyMode = False     'コピーする範囲(←コピー元にある)を示す、点滅した点線の範囲が解除される
    
    'コミッション列を削除
    Columns("CB:CB").Delete Shift:=xlToLeft '配布ブックにはコミッションを含めない
    
    shtOwr.Rows("1:1").AutoFilter
    shtOwr.Range("K1").Select
    
    '↓↓進捗報告ブック
    shtPrg.Activate
    ActiveSheet.AutoFilterMode = False                  'フィルタを初期化 ⇒ 上で設定された条件を消す
    shtPrg.Range("$B$1:$BS$43").AutoFilter Field:=8  'AutoFilterの引数Criteria1を指定しない ⇒ フィルタを絞らない(全選択)
    shtPrg.Range("K1").Select
    
'    '↓↓MR別ブックファイル
'    shtOwr.Activate
    
    Exit Sub
    
Err1:
    MsgBox "MR別ブック作成()" & vbCrLf & _
           "エラー番号:" & Err.Number & vbCrLf & _
           "エラーの種類:" & Err.Description, vbExclamation
           
End Sub

Sub データ範囲セット(strSheetName As String, strRangeName As String, intTrgRow As Integer, intTrgColumn As Integer)

    'A1から右下端のデータ範囲に名前"Data_Range"をセットする
    '引数intTrgRow    : intTrgRow列目の列は、最下端のセルまで値が存在する列を選ぶ
    '引数intTrgColumn : intTrgColumn行目の行は、最右端のセルまで値が存在する列を選ぶ(通常は１行目のヘッダー行)
    
    Dim strEndRow As Integer        '最終行の行番号
    Dim strEndColumn As Integer     '最終列の列番号
    
    'On Error GoTo Err1
    
    Application.ScreenUpdating = False
    
    '↓↓引数strSheetNameで指定されたシート(担当者など)
    Sheets(strSheetName).Activate
    Sheets(strSheetName).Range("A1").Select     'データ範囲外のセルにカーソルがあると、処理の挙動がおかしくなる　⇒　A1にカーソルをおいておけば問題無い
    strEndRow = Cells(Rows.Count, intTrgRow).End(xlUp).Row                      'intTrgRow列目の列の一番下のセル Cells(Rows.Count, intTrgRow)
    strEndColumn = Cells(intTrgColumn, Columns.Count).End(xlToLeft).Column      'intTrgColumn行目の行の一番右のセル Cells(intTrgColumn, Columns.Count)
    
    'A1から右下端のデータ範囲に名前"Data_Range"をセットする
    Sheets(strSheetName).Names.Add Name:=strRangeName, RefersToR1C1:="=" & strSheetName & "!R1C1:R" & strEndRow & "C" & strEndColumn
    Sheets(strSheetName).Names(strRangeName).Comment = ""
    
'        'このプロシージャ単体で使う場合はコメントアウト
'        'カーソル位置を初期化
'        Range("A1").Select
'        ActiveWindow.ScrollRow = 1
'        ActiveWindow.ScrollColumn = 1
    
    Exit Sub
    
Err1:
    MsgBox "データ範囲セット()" & vbCrLf & _
           "エラー番号:" & Err.Number & vbCrLf & _
           "エラーの種類:" & Err.Description, vbExclamation
           
End Sub

Sub AWK進捗報告前処理()

    Dim wbkPrg As Workbook
    Dim objCnfWbk As Workbook
    Dim shtPrg As Worksheet
    Dim strEndRow As Integer    '最終行の行番号
    Dim strIspClm As String     '列名のアルファベット
    
    'On Error GoTo Err1
    
    'Application.Cursor = xlWait                      'マウスカーソルを砂時計ににする
    Application.ScreenUpdating = False
    
    '以降の処理でアクティブブックが変わる前に、
    'プロシージャ呼出し時のアクティブブックとアクティブシートを取得しておく
    Set wbkPrg = ActiveWorkbook
    Set shtPrg = ActiveSheet
    
    '並び替え　ISP契約番号(昇順)
    'Order ⇒ xlAscending(昇順),xlDescending(降順)
    strEndRow = Cells(Rows.Count, 1).End(xlUp).Row  'A列目の一番下のセル Cells(Rows.Count, 1)
    shtPrg.Sort.SortFields.Clear
    
    'ISP契約番号の列名アルファベットを取得する
    strIspClm = Cells.Find(What:="ISP契約番号", After:=ActiveCell, LookIn:=xlFormulas, _
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
    
    '前回処理で列が生成されていれば削除
    If shtPrg.Range("A1").FormulaR1C1 = "集計日" _
        And shtPrg.Range("I1").FormulaR1C1 = "担当者" _
        And shtPrg.Range("J1").FormulaR1C1 = "申込月" _
        And shtPrg.Range("K1").FormulaR1C1 = "フェーズ" _
        And shtPrg.Range("CB1").FormulaR1C1 = "コミッション" Then
        Range("A:A,I:I,J:J,K:K,CB:CB").Delete Shift:=xlToLeft
    End If
    
    '追加フィールド列挿入
    shtPrg.Select
    shtPrg.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    shtPrg.Range("A1").FormulaR1C1 = "集計日"
    shtPrg.Columns("I:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    shtPrg.Range("I1").FormulaR1C1 = "担当者"
    shtPrg.Columns("J:J").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    shtPrg.Range("J1").FormulaR1C1 = "申込月"
    shtPrg.Columns("K:K").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    shtPrg.Range("K1").FormulaR1C1 = "フェーズ"
    shtPrg.Columns("CB:CB").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    shtPrg.Columns("CB:CB").NumberFormatLocal = "\#,##0;\-#,##0"
    shtPrg.Range("CB1").FormulaR1C1 = "コミッション"
    
    'ブック 開く
    '「取り込み元」
    Set objCnfWbk = Workbooks.Open(wbkPrg.Path & "\NURO進捗報告設定.xls", Password:="0911")
    
    'シート 取り込み
    '「担当者」
    If ExistSheet("担当者", wbkPrg) Then '既に取り込んでいれば削除
        Application.DisplayAlerts = False
        wbkPrg.Worksheets("担当者").Delete
        Application.DisplayAlerts = True
    End If
    objCnfWbk.Worksheets("担当者").Copy After:=wbkPrg.Worksheets(wbkPrg.Worksheets.Count)
    '「キャンペーン」
    If ExistSheet("キャンペーン", wbkPrg) Then   '既に取り込んでいれば削除
        Application.DisplayAlerts = False
        wbkPrg.Worksheets("キャンペーン").Delete
        Application.DisplayAlerts = True
    End If
    objCnfWbk.Worksheets("キャンペーン").Copy After:=wbkPrg.Worksheets(wbkPrg.Worksheets.Count)
    '「配布」
    If ExistSheet("配布", wbkPrg) Then '既に取り込んでいれば削除
        Application.DisplayAlerts = False
        wbkPrg.Worksheets("配布").Delete
        Application.DisplayAlerts = True
    End If
    objCnfWbk.Worksheets("配布").Copy After:=wbkPrg.Worksheets(wbkPrg.Worksheets.Count)
    
    '設定ファイルを保存しないで閉じる
    objCnfWbk.Close saveChanges:=False
    
    'プロシージャ呼出し時のアクティブブックに戻る
    shtPrg.Activate
    shtPrg.Range("K1").Select
    
    Exit Sub
    
Err1:
    MsgBox "AWK進捗報告前処理()" & vbCrLf & _
           "エラー番号:" & Err.Number & vbCrLf & _
           "エラーの種類:" & Err.Description, vbExclamation
           
End Sub

Function SQL生成() As String

    Dim shtPrg As Worksheet
    
    'On Error GoTo Err1
    
    Application.ScreenUpdating = False
    
    '以降の処理でアクティブブックが変わる前に、
    'プロシージャ呼出し時のアクティブシートを取得しておく
    Set shtPrg = ActiveWorkbook.ActiveSheet
    
    '※予め、データ範囲に名前を付けておく
    '※ブック上のデータ範囲ではなく、シート上のデータ範囲に名前を付ける
    Call データ範囲セット("SP_WORK", "Data_Range", 8, 1)   '空白が無い列と行の、番号をセット　⇒　８列目：ISP契約番号、１行目：列名ヘッダー行
    Call データ範囲セット("担当者", "Owner_Range", 1, 1)   '空白が無い列と行の、番号をセット　⇒　１列目：ISP契約番号、１行目：列名ヘッダー行
    Call データ範囲セット("キャンペーン", "Campaign_Range", 1, 1)   '空白が無い列と行の、番号をセット　⇒　１列目：ISP契約番号、１行目：列名ヘッダー行
    
    SQL生成 = "select " & _
                    "[SP_WORK$Data_Range].[申込日], " & _
                    "[担当者$Owner_Range].[担当者], " & _
                    "[SP_WORK$Data_Range].[会員氏名], " & _
                    "[SP_WORK$Data_Range].[So-net工事予定日], " & _
                    "[SP_WORK$Data_Range].[NTT工事予定日], " & _
                    "[SP_WORK$Data_Range].[So-net工事日], " & _
                    "[SP_WORK$Data_Range].[NTT工事日], " & _
                    "[SP_WORK$Data_Range].[NURO光回線開通処理日], " & _
                    "[SP_WORK$Data_Range].[決済情報確定日], " & _
                    "[SP_WORK$Data_Range].[キャンセル日], " & _
                    "[SP_WORK$Data_Range].[ご連絡先電話番号], " & _
                    "[キャンペーン$Campaign_Range].[コミッション] " & _
              "from  [SP_WORK$Data_Range], " & _
                    "[担当者$Owner_Range], " & _
                    "[キャンペーン$Campaign_Range] " & _
              "Where [SP_WORK$Data_Range].[ISP契約番号] = [担当者$Owner_Range].[ISP契約番号] " & _
              "And   [SP_WORK$Data_Range].[代理店コード] = [キャンペーン$Campaign_Range].[キャンペーンコード] " & _
              "Order by [SP_WORK$Data_Range].[ISP契約番号] "
    
    'プロシージャ呼出し時のアクティブブックに戻る
    shtPrg.Activate
    shtPrg.Range("K1").Select
    
    Exit Function
    
Err1:
    MsgBox "SQL生成()" & vbCrLf & _
           "エラー番号:" & Err.Number & vbCrLf & _
           "エラーの種類:" & Err.Description, vbExclamation
           
End Function

Sub AWK進捗報告()

    Dim objCn As New ADODB.Connection
    Dim objRs  As ADODB.Recordset
    Dim strSQL As String
    Dim strBuf As String
    Dim intRow As Integer
    Dim dtm集計日 As Date
    Dim strフェーズ As String
    Dim wbkPrg As Workbook
    Dim shtPrg As Worksheet
    
    'On Error GoTo Err1
    
    '以降の処理でアクティブブックが変わる前に、
    'プロシージャ呼出し時のアクティブブックとアクティブシートを取得しておく
    Set wbkPrg = ActiveWorkbook
    Set shtPrg = ActiveSheet
    
    Call AWK進捗報告前処理
    
    With objCn
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & wbkPrg.Path & "\" & wbkPrg.Name & ";" & _
         "Extended Properties=Excel 8.0;"
        .Open
    End With
    
    'SQL生成() ⇒ データ範囲設定も同時に行う
    strSQL = SQL生成
    Set objRs = objCn.Execute(strSQL)
    
'                            Debug.Print "レコード数" & objRs.Fields.Count
'
'                            '列番列名リストをダンプ(イミディエイトウィンドウへ)
'                            Dim i
'                            i = 0
'                            For i = 0 To objRs.Fields.Count - 1
'                                Debug.Print i & " " & objRs.Fields(i).Name
'                            Next
'
'                            '列名ヘッダーをダンプ(イミディエイトウィンドウへ)
'                            strBuf = objRs.Fields(0).Name _
'                             & "," & objRs.Fields(1).Name _
'                             & ",申込月" _
'                             & ",フェーズ" _
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
    
    intRow = 1   '行数
    dtm集計日 = CDate(Format(Mid(wbkPrg.Name, 9, 8), "@@@@/@@/@@"))
    Do While objRs.EOF = False
    
        'フェーズを取得
        strフェーズ = getフェーズ( _
                                    CStr(dtm集計日), _
                                    IIf(IsNull(objRs![申込日]), "", objRs![申込日]), _
                                    IIf(IsNull(objRs![So-net工事予定日]), "", objRs![So-net工事予定日]), _
                                    IIf(IsNull(objRs![NTT工事予定日]), "", objRs![NTT工事予定日]), _
                                    IIf(IsNull(objRs![So-net工事日]), "", objRs![So-net工事日]), _
                                    IIf(IsNull(objRs![NTT工事日]), "", objRs![NTT工事日]), _
                                    IIf(IsNull(objRs![NURO光回線開通処理日]), "", objRs![NURO光回線開通処理日]), _
                                    IIf(IsNull(objRs![決済情報確定日]), "", objRs![決済情報確定日]), _
                                    IIf(IsNull(objRs![キャンセル日]), "", objRs![キャンセル日]) _
                                )
        '申込月を取得
        str申込月 = Chk月(objRs![申込日], CStr(dtm集計日))
        
'                            'データをダンプ(イミディエイトウィンドウ)
'                            strBuf = objRs![申込日] _
'                            & "," & objRs![担当者] _
'                            & "," & str申込月 & "申込" _
'                            & "," & strフェーズ _
'                            & "," & objRs![会員氏名] _
'                            & "," & objRs![So-net工事予定日] _
'                            & "," & objRs![NTT工事予定日] _
'                            & "," & objRs![So-net工事日] _
'                            & "," & objRs![NTT工事日] _
'                            & "," & objRs![NURO光回線開通処理日] _
'                            & "," & objRs![決済情報確定日] _
'                            & "," & objRs![キャンセル日] _
'                            & "," & IIf(strMode = "admin", objRs![コミッション], "")
'                            Debug.Print strBuf
        
        'シートへ書き込む
        shtPrg.Range("A" & intRow + 1) = CStr(dtm集計日)
        shtPrg.Range("I" & intRow + 1) = objRs![担当者]
        shtPrg.Range("J" & intRow + 1) = str申込月 & "申込"
        shtPrg.Range("K" & intRow + 1) = strフェーズ
        If Not IsNull(objRs![ご連絡先電話番号]) Then
            If InStr(objRs![ご連絡先電話番号], "-") = 0 Then
                shtPrg.Range("N" & intRow + 1) = "0" & Mid(objRs![ご連絡先電話番号], 1, 1) & "-" & Mid(objRs![ご連絡先電話番号], 2, 4) & "-" & Mid(objRs![ご連絡先電話番号], 6, 4)
            End If
        End If
        shtPrg.Range("CB" & intRow + 1) = objRs![コミッション]
        
        intRow = intRow + 1
        objRs.MoveNext  '次のレコードへ移動
    Loop
    
    objRs.Close
    Set objRs = Nothing
    
    
    Call AWK進捗報告後処理
    
    '配布処理
    '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
    
    Call データ範囲セット("配布", "Dist_Range", 1, 1)
    
    'プロシージャ呼出し時のアクティブブックに戻る
    shtPrg.Activate
    shtPrg.Range("K1").Select
    
    strSQL = "select " & _
                    "[配布$Dist_Range].[配布者], " & _
                    "[配布$Dist_Range].[配布条件] " & _
              "from  [配布$Dist_Range]"
    
    Set objRs = objCn.Execute(strSQL)
    Do While objRs.EOF = False
    
        Call 配布ブック作成(objRs![配布者], objRs![配布条件])
        
        objRs.MoveNext  '次のレコードへ移動
    Loop
    
    If ExistSheet("配布") Then
        Application.DisplayAlerts = False
        Worksheets("配布").Delete
        Application.DisplayAlerts = True
    End If
    
    '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
    
    objRs.Close
    objCn.Close
    
    Set objRs = Nothing
    Set objCn = Nothing
    
    Exit Sub
    
Err1:
    MsgBox "AWK進捗報告()" & vbCrLf & _
           "エラー番号:" & Err.Number & vbCrLf & _
           "エラーの種類:" & Err.Description, vbExclamation
           
End Sub

Sub AWK進捗報告後処理()

    Dim shtPrg As Worksheet
    
    'On Error GoTo Err1
    
    Application.ScreenUpdating = False
    
    Set shtPrg = ActiveSheet
    
    Call フェーズ_条件付書式_色設定
    
    'シートが存在していれば削除
    If ExistSheet("担当者") Then
        Application.DisplayAlerts = False
        Worksheets("担当者").Delete
        Application.DisplayAlerts = True
    End If
    If ExistSheet("キャンペーン") Then
        Application.DisplayAlerts = False
        Worksheets("キャンペーン").Delete
        Application.DisplayAlerts = True
    End If
    
    '"SP_WORK"を選択
    shtPrg.Select
    
    'オートフィルタ設定
    If ActiveSheet.AutoFilterMode = False Then
        shtPrg.Rows("1:1").AutoFilter
    End If
    
    '列幅を最適化
    shtPrg.Columns("A:A").EntireColumn.AutoFit   '集計日(追加した列)
    shtPrg.Columns("H:H").EntireColumn.AutoFit   'ISP契約番号
    shtPrg.Columns("I:K").EntireColumn.AutoFit   '担当者,申込月,フェーズ追加した列
    shtPrg.Columns("N:N").EntireColumn.AutoFit   'ご連絡先電話番号
    shtPrg.Columns("S:S").EntireColumn.AutoFit   '申込日
    shtPrg.Columns("Y:AD").EntireColumn.AutoFit   'So-net工事予定日,NTT工事予定日,So-net工事日,NTT工事日,NURO光回線開通処理日,キャンセル日
    shtPrg.Columns("AF:AF").EntireColumn.AutoFit   '申込日
    
    '処理終了後カーソル位置
    shtPrg.Range("K1").Select
    
    Application.Cursor = xlDefault                      'マウスカーソルを砂時計から元ににする
    
    Exit Sub
    
Err1:
    MsgBox "AWK進捗報告後処理()" & vbCrLf & _
           "エラー番号:" & Err.Number & vbCrLf & _
           "エラーの種類:" & Err.Description, vbExclamation
           
End Sub

Sub フェーズ_条件付書式_色設定()

    'On Error GoTo Err1
    
    Sheets("SP_WORK").Activate
    
    Columns("K:K").Select
    
                '＊＊＊＊　条件付き書式
                'FormatConditions.Add　条件付き書式を追加
                'Type:=xlTextString ⇒ 特定の文字列
                'TextOperator:=xlContains ⇒ 次の値を含む セルの値次の値に等しい
                'String:=　←で指定
                
                'Type:=xlCellValue ⇒ セルの値
                'Operator:=xlEqual ⇒ 次の値に等しい
                'Formula1:=　←で指定
                
                '優先順位を１位にして条件付き書式を追加する
                'Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                
                '条件付き書式の『条件を満たす場合は停止』　⇒　そこから「下の」(＝↑既に設定した)条件を調べるかどうかを設定
                'Selection.FormatConditions(1).StopIfTrue = True
    
    '予定
    Selection.FormatConditions.Add Type:=xlTextString, String:="未定", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.シーグリーン  '50
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    '予定
    Selection.FormatConditions.Add Type:=xlTextString, String:="予定", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.ライム  '43
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    'S完了
    Selection.FormatConditions.Add Type:=xlTextString, String:="S完了", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.ゴールド '44
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    'N完了
    Selection.FormatConditions.Add Type:=xlTextString, String:="N完了", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.薄いオレンジ '45
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    '開通
    Selection.FormatConditions.Add Type:=xlTextString, String:="開通", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.オレンジ '46
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    '確定
    Selection.FormatConditions.Add Type:=xlTextString, String:="確定", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ColorIndex = ClrIdx.白 '2
    End With
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.赤 '3
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    '来月
    Selection.FormatConditions.Add Type:=xlTextString, String:="来月", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.水色 '8
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    'CXL
    Selection.FormatConditions.Add Type:=xlTextString, String:="CXL", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ColorIndex = ClrIdx.白 '2
    End With
    With Selection.FormatConditions(1).Interior
        .ColorIndex = ClrIdx.G紫 '29
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    Exit Sub
    
Err1:
    MsgBox "フェーズ_条件付書式_色設定()" & vbCrLf & _
           "エラー番号:" & Err.Number & vbCrLf & _
           "エラーの種類:" & Err.Description, vbExclamation
           
End Sub

Function getフェーズ(str集計日 As String, str申込日 As String, strS予定日 As String, strN予定日 As String, _
                        strS工事日 As String, strN工事日 As String, str開通日 As String, str決済確定日 As String, strCXL日 As String) As String
    
    Dim flg_申込日 As String
    Dim flg_SON予定日 As String
    Dim flg_N予定日 As String
    Dim flg_SON工事日 As String
    Dim flg_N工事日 As String
    Dim flg_開通日 As String
    Dim flg_決済確定日 As String
    Dim flg_CXL日 As String
    
    'On Error GoTo Err1
    
    Application.ScreenUpdating = False
    
    flg_申込日 = Chk月(str申込日, str集計日)
    flg_SON予定日 = Chk月(strS予定日, str集計日)
    flg_N予定日 = Chk月(strN予定日, str集計日)
    flg_SON工事日 = Chk月(strS工事日, str集計日)
    flg_N工事日 = Chk月(strN工事日, str集計日)
    flg_開通日 = Chk月(str開通日, str集計日)
    flg_決済確定日 = Chk月(str決済確定日, str集計日)
    flg_CXL日 = Chk月(strCXL日, str集計日)
    
    Select Case flg_申込日                                              '↓↓フェーズの集計対象：申込日が前々月以降
    Case "当月", "前月", "前々月"                                       '↓↓CXLの集計対象：申込日が前々月以降
        Select Case flg_CXL日
        Case "当月"
            getフェーズ = "CXL" '実質は、当月CXL
        Case "空白"                                                     '↓↓開通または確定の集計対象：CXL日が空白
            Select Case flg_開通日
            Case "当月"
                Select Case flg_決済確定日
                Case "空白"
                    getフェーズ = "当月開通"
                Case "前々々月以前", "前々月", "前月", "当月", "来月", "再来月以降"
                    getフェーズ = "当月確定"
                Case Else
                    getフェーズ = ""
                End Select
                
            Case "前月"
                Select Case flg_決済確定日
                Case "空白"
                    getフェーズ = "前月開通"
                Case "前々々月以前", "前々月", "前月", "当月", "来月", "再来月以降"
                    'getフェーズ = "前月確定"
                    getフェーズ = ""               '集計対象外：前月確定
                Case Else
                    getフェーズ = ""
                End Select
                
            Case "前々月"
                Select Case flg_決済確定日
                Case "空白"
                    getフェーズ = "前々月開通"
                Case "前々々月以前", "前々月", "前月", "当月", "来月", "再来月以降"
                    'getフェーズ = "前々月確定"
                    getフェーズ = ""               '集計対象外：前々月確定
                Case Else
                    getフェーズ = ""
                End Select
                
            Case "空白"                                             '↓↓N完了の集計対象：開通が空白
                Select Case flg_N工事日
                Case "当月"
                    getフェーズ = "当月N完了"
                Case "前月"
                    getフェーズ = "前月N完了"
                Case "前々月"
                    getフェーズ = "前々月N完了"
                Case "空白"                                         '↓↓S完了の集計対象：開通が空白
                    Select Case flg_SON工事日
                    Case "当月"
                        getフェーズ = "当月S完了"
                    Case "当月"
                        getフェーズ = "前月S完了"
                    Case "当月"
                        getフェーズ = "前々月S完了"
                    Case "空白"                                     '↓↓SN予定,未定の集計対象：開通が空白
                        Select Case flg_N予定日
                        Case "来月"
                            Select Case flg_SON予定日
                            Case "来月"
                                getフェーズ = "来月SN予定"          '来月S来月N予定
                            Case "当月"
                                getフェーズ = "来月SN予定"          '当月S来月N予定
                            Case "前月"
                                getフェーズ = "来月SN予定"          '前月S来月N予定
                            Case "前々月"
                                getフェーズ = "来月SN予定"          '前々月S来月N予定
                            Case "空白"
                                getフェーズ = "来月Nのみ予定"       '未定S来月N予定
                            End Select
                            
                        Case "当月"
                            Select Case flg_SON予定日
                            Case "来月"
                                getフェーズ = "来月SN予定"          '来月S当月N予定
                            Case "当月"
                                getフェーズ = "当月SN予定"          '当月S当月N予定
                            Case "前月"
                                getフェーズ = "当月SN予定"          '前月S当月N予定
                            Case "前々月"
                                getフェーズ = "当月SN予定"          '前々月S当月N予定
                            Case "空白"
                                getフェーズ = "当月Nのみ予定"       '未定S当月N予定
                            End Select
                            
                        Case "前月"
                            Select Case flg_SON予定日
                            Case "来月"
                                getフェーズ = "来月SN予定"          '来月S前月N予定
                            Case "当月"
                                getフェーズ = "当月SN予定"          '当月S前月N予定
                            Case "前月"
                                getフェーズ = "前月SN予定"          '前月SN予定
                            Case "前々月"
                                getフェーズ = "前月SN予定"          '前々月S前月N予定
                            Case "空白"
                                getフェーズ = "前月Nのみ予定"       '未定S前月N予定
                            End Select
                            
                        Case "前々月"
                            Select Case flg_SON予定日
                            Case "来月"
                                getフェーズ = "来月SN予定"          '来月S前々月N予定
                            Case "当月"
                                getフェーズ = "当月SN予定"          '当月S前々月N予定
                            Case "前月"
                                getフェーズ = "前月SN予定"          '前月S前々月N予定
                            Case "前々月"
                                getフェーズ = "前々月SN予定"        '前々月S前々月N予定
                            Case "空白"
                                getフェーズ = "前々月Nのみ予定"     '未定S前々月N予定
                            End Select
                            
                        Case "空白"
                            Select Case flg_SON予定日
                            Case "来月"
                                getフェーズ = "来月Sのみ予定"       '来月S未定N予定
                            Case "当月"
                                getフェーズ = "当月Sのみ予定"       '当月S未定N予定
                            Case "前月"
                                getフェーズ = "前月Sのみ予定"       '前月S未定N予定
                            Case "前月"
                                getフェーズ = "前々月Sのみ予定"         '前々月S未定N予定
                            Case "空白"
                                getフェーズ = "未定"                '未定S未定N予定
                            End Select
                            
                        End Select
                    Case Else
                        getフェーズ = ""    '集計対象外：S工事日が、不正値
                    End Select
                Case Else
                    getフェーズ = ""    '集計対象外：N工事日が、不正値
                End Select
            Case Else
                getフェーズ = ""    '集計対象外：開通日が、不正値
            End Select
        Case Else
            getフェーズ = ""    '集計対象外：CXL日が、前月以前または不正値
        End Select
    Case Else
        getフェーズ = ""    '集計対象外：申込日が、前々月以前または不正値
    End Select
    
    Exit Function
    
Err1:
    MsgBox "getフェーズ()" & vbCrLf & _
           "エラー番号:" & Err.Number & vbCrLf & _
           "エラーの種類:" & Err.Description, vbExclamation
           
End Function

Function Chk月(strTrgDate As String, strAggDate As String) As String

    'On Error GoTo Err1
    
    Application.ScreenUpdating = False
    
    'strTrgDate : 評価日付
    'strAggDate : 集計日付
    
    If strTrgDate = "" Then '評価日付が、空白なら。。。
    
        Chk月 = "空白"
        
    '評価日付が、空白でないなら。。。
    Else
        
        '(参考)　月初取得 DateSerial(Year("yyyy/mm/dd"), Month("yyyy/mm/dd"), 1)
        '　　　  月末取得 DateSerial(Year("yyyy/mm/dd"), Month("yyyy/mm/dd"), 0)
        
        Select Case CDate(strTrgDate)   '評価日付が、以下の範囲内なら。。。
        
            '再来月月初 以降
            Case Is >= DateSerial(Year(strAggDate), Month(strAggDate) + 2, 1)
                Chk月 = "再来月以降"
                
            '来月月初 から 来月月末 の間
            Case DateSerial(Year(strAggDate), Month(strAggDate) + 1, 1) To DateSerial(Year(strAggDate), Month(strAggDate) + 2, 0)
                Chk月 = "来月"
                
            '当月月初 から 当月月末 の間
            Case DateSerial(Year(strAggDate), Month(strAggDate) + 0, 1) To DateSerial(Year(strAggDate), Month(strAggDate) + 1, 0)
                Chk月 = "当月"
                
            '前月月初 から 前月月末 の間
            Case DateSerial(Year(strAggDate), Month(strAggDate) - 1, 1) To DateSerial(Year(strAggDate), Month(strAggDate) + 0, 0)
                Chk月 = "前月"
                
            '前々月月初 から 前々月月末 の間
            Case DateSerial(Year(strAggDate), Month(strAggDate) - 2, 1) To DateSerial(Year(strAggDate), Month(strAggDate) - 1, 0)
                Chk月 = "前々月"
                
            '前々月月初 より過去
            Case Is < DateSerial(Year(strAggDate), Month(strAggDate) - 2, 1)
                Chk月 = "前々々月以前"
                
        End Select
        
    End If
    
    Exit Function
    
Err1:
    MsgBox "Chk月()" & vbCrLf & _
           "エラー番号:" & Err.Number & vbCrLf & _
           "エラーの種類:" & Err.Description, vbExclamation
           
End Function

Sub 基本形_ExcelへODBC接続()

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
    
    'データ範囲に予め名前を付けておく(ブックではなく、シート上の範囲にする)
    'SQLのテーブル名は、[シート名$データ範囲名]のフォーマットで指定する
    
    'strSQL = "select * from [" & ActiveSheet.Name & "$Data_Range]"
    strSQL = "select * from [" & ActiveSheet.Name & "$SP_WORK]"
    Set objRs = objCn.Execute(strSQL)
    
        '    '列番列名ダンプ
        '    Dim i
        '    i = 0
        '    For i = 0 To Cells(1, Columns.Count).End(xlToLeft).Column - 2
        '        Debug.Print i & " " & objRs.Fields(i).Name
        '    Next
    
    '列名ダンプ
    strBuf = objRs.Fields(10).Name & "," & objRs.Fields(23).Name & "," & objRs.Fields(24).Name & "," & objRs.Fields(25).Name & "," & objRs.Fields(26).Name & "," & objRs.Fields(27).Name & "," & objRs.Fields(28).Name
    Debug.Print strBuf
    'ActiveSheet.Range("A40:I40") = Split(strBuf, ",")
    
    'データダンプ
    Do While objRs.EOF = False
        Debug.Print objRs!会員氏名 & ", " & objRs![S工事予定日] & ", " & objRs![N工事予定日] & ", " & objRs![S工事日] & ", " & objRs![N工事日] & ", " & objRs![NURO光回線開通処理日] & ", " & objRs![キャンセル日]
        objRs.MoveNext  '次のレコードへ移動
    Loop
    
    objRs.Close
    objCn.Close
    
    Set objRs = Nothing
    Set objCn = Nothing
    
    Exit Sub
    
Err1:
    MsgBox "基本形_ExcelへODBC接続()" & vbCrLf & _
           "エラー番号:" & Err.Number & vbCrLf & _
           "エラーの種類:" & Err.Description, vbExclamation
           
End Sub

'ワークシートが存在するか調べる関数
Function ExistSheet(strSheetName As String, Optional objWbk As Variant) As Boolean

    Dim objSheet As Object
    
    'On Error GoTo Err1
      
    ExistSheet = False
    'IsMissingの引数はVariant型である必要あり
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
           "エラー番号:" & Err.Number & vbCrLf & _
           "エラーの種類:" & Err.Description, vbExclamation
           
End Function



