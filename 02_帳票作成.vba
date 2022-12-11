' ----------------------------------------------------------------------
'   ファイル名：帳票作成
'   動作環境：	Excel
'   機能説明：  売上データと帳票テンプレートが格納されたフォルダを指定
'				2つのファイルを元に帳票を作成
' ----------------------------------------------------------------------

'定数定義
'表示メッセージ
Const MSG_ST As String = "処理を開始します。フォルダを指定して下さい。"
Const DIR_TITLE As String = "対象フォルダを選択"
Const MSG_END As String = "処理が終了しました"
Const MSG_NF As String = "が見つかりません"
'ファイル名
Const FN_SALES As String = "01_売上データ.xlsx"
Const FN_TEMP As String = "09_TEMPLATE.xlsx"
Const FN_OUTPUT As String = "サンプル帳票_?????時点.xlsx"
'数式
Const FML_ST As String = "=SUBTOTAL(9, ?????)"
'文字列
Const SHEET_TITLE As String = "支店別売上一覧"
Const SUM_BR As String = "支店計"
Const SUM_TEAM As String = "チーム計"
Const TEMP_WORD As String = "?????"
'書式
Const COMMA As String = "#,##0"
Const THOU As String = "#,##0,"
Const STR_YMD As String = "YYYY年MM月DD日"
Const INT_YMD As String = "yymmdd"
Const YM As String = "yyyy年mm月"
'ヘッダー
Const HEAD_MO As Long = 5
Const HEAD_CNT As Long = 13
Const LAST_MO As String = "以降（13ヶ月目以降）"
Const THIS_MO As String = "（当月）"
Const MO_UNIT As String = "ヶ月目）"
'その他数値
Const TITLE_POS As String = "A1"
Const COPY_COEF As Long = 3
Const PASTE_COEF As Long = 4
Const ELEM_CNT As Long = 2
Const ITEM_CNT As Long = 14
Const ITEM_END As Long = 62
Const DATA_START As Long = 5
Const GT As Long = 7
Const BEGIN As Long = 9
Const GREY As Long = 15

'列挙型定義
'出力ファイルの列
Enum OUTPUT
    CODE = 2
    BRANCH = 3
    TEAM = 4
    BOOK = 6
End Enum

'処理フラグ
Enum FLG
    BR_ST = 1
    TEAM_ST = 2
End Enum

'ユーザー定義型宣言
'作業ファイル
Type Booklist
    Data As Workbook
    Main As Workbook
End Type

'日付
Type Datelist
    BaseDate As Date
    BaseDateStr As String
End Type

' ----------------------------------------------------------------------
'   関数名：MainProc（主処理）
'   引数：  なし
'   戻り値：なし
' ----------------------------------------------------------------------
Sub MainProc()
    '変数宣言
    Dim TargetBook As Booklist
    Dim FolderPath As String
    Dim CopyFile As String
    Dim ToCopyPath As String
    Dim TargetRng As Range
    
    '実行確認
    If MsgBox(MSG_ST, vbOKCancel + vbInformation) = vbCancel Then Exit Sub
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = DIR_TITLE                      'タイトルの指定
        .InitialFileName = ThisWorkbook.Path    '初期表示フォルダの指定
        If .Show = True Then                    'ダイアログを表示して戻り値を判定
            FolderPath = .SelectedItems(1)      'フォルダのパスを取得
        Else
            Exit Sub
        End If
    End With
    
    '初期処理
    If InitProc(TargetBook, FolderPath) = False Then Exit Sub
    
    'テンプレートファイルを別名で保存
    CopyFile = Replace(FN_OUTPUT, TEMP_WORD, Format(Date, INT_YMD))
    ToCopyPath = ThisWorkbook.Path & "\" & CopyFile
    FileCopy FolderPath & "\" & FN_TEMP, ToCopyPath
    
    '別名保存したファイルを開く
    Set TargetBook.Main = Workbooks.Open(ToCopyPath)
    
    '売上データのコピーと貼り付け
    Call GetSalesData(TargetBook)
    
    '小計行作成
    With ActiveSheet
        Set TargetRng = .Range(.Cells(BEGIN, OUTPUT.CODE), .Cells(Rows.Count, OUTPUT.CODE).End(xlUp))
    End With
    Call MakeSubTotal(TargetRng, FLG.BR_ST)
    
    '書式設定
    Call EditFormat(TargetBook.Main.Sheets(1))
    
    'ファイルを閉じる
    With TargetBook.Main
        .Sheets(1).Activate
        .Sheets(1).Range(TITLE_POS).Select
        .Close SaveChanges:=True
    End With
    Set TargetBook.Main = Nothing
    
    '終了処理
    Call EndProc
End Sub

' ----------------------------------------------------------------------
'   関数名：InitProc（初期処理）
'   引数1： TargetBook（エクセルファイルリスト）
'   引数2： FolderPath（ユーザーが指定したフォルダのパス）
'   戻り値：Bollean（正常=True/異常=False）
' ----------------------------------------------------------------------
Function InitProc(TargetBook As Booklist, FolderPath As String) As Boolean
    
    With Application
        '画面更新停止
        .ScreenUpdating = False
        '確認メッセージ非表示
        .DisplayAlerts = False
    End With
    
    'カレントディレクトリ変更
    ChDir ThisWorkbook.Path

    '01_売上データ.xlsxファイルを開く
    If (Dir(FolderPath & "\" & FN_SALES) = "") Then
        MsgBox FN_SALES & MSG_NF, vbCritical
        GoTo ErrProc
    End If
    Set TargetBook.Data = Workbooks.Open(FolderPath & "\" & FN_SALES)
       
    '09_TEMPLATE.xlsxが存在することを確認
    If (Dir(FolderPath & "\" & FN_TEMP) = "") Then
        MsgBox FN_TEMP & MSG_NF, vbCritical
        TargetBook.Data.Close
        GoTo ErrProc
    End If
    
    '正常終了
    InitProc = True
    Exit Function
    
ErrProc:  'エラー処理
    Call EndProc
    InitProc = False
    Exit Function
End Function

' ----------------------------------------------------------------------
'   関数名：GetSalesData（売上データ取得処理）
'   引数：  TargetBook（エクセルファイルリスト）
'   戻り値：なし
' ----------------------------------------------------------------------
Sub GetSalesData(TargetBook As Booklist)
    '変数宣言
    Dim TargetSheet As Worksheet
    Dim CopyRng As Range
    Dim CopyStPos As Long
    Dim PasteStPos As Long
    Dim MainCnt As Long
    Dim SubCnt As Long
    Dim ColChar As String
    Dim TargetCol As Long
    Dim StRngStr As String
    
    'ヘッダーを設定
    Set TargetSheet = TargetBook.Main.Sheets(1)
    Call EditHeader(TargetSheet)
    
    '売上データの範囲を取得
    With TargetBook.Data.Sheets(1).UsedRange
        Set CopyRng = .Offset(1).Resize(.Rows.Count - 1)
    End With
    
    '出力ファイルに見出しを転記
    CopyRng.Resize(, 4).Copy
    TargetSheet.Cells(BEGIN, OUTPUT.CODE).PasteSpecial Paste:=xlPasteValues
    '罫線設定
    Call EditLine(Selection)
    Selection.End(xlDown).Offset(1).Resize(, 4).Borders(xlEdgeTop).LineStyle = xlContinuous
    
    '出力ファイルに売上データを転記
    For MainCnt = 0 To ITEM_CNT
        CopyStPos = DATA_START + (COPY_COEF * MainCnt) - 1
        PasteStPos = OUTPUT.BOOK + (PASTE_COEF * MainCnt)
        
        'コピーして数値形式で貼り付け
        CopyRng.Offset(, CopyStPos).Resize(, 3).Copy
        TargetSheet.Cells(BEGIN, PasteStPos).PasteSpecial Paste:=xlPasteValues
        
        '罫線設定
        Call EditLine(Selection)
        Selection.End(xlDown).Offset(1).Resize(, 3).Borders(xlEdgeTop).LineStyle = xlContinuous
        '合計行に数式を入力
        For SubCnt = 0 To ELEM_CNT
            TargetCol = PasteStPos + SubCnt
            ColChar = Replace(TargetSheet.Cells(1, TargetCol).Address(True, False), "$1", "")
            With TargetSheet.Cells(BEGIN, TargetCol)
                StRngStr = ColChar & .Row & ":" & ColChar & .End(xlDown).Row
            End With
            TargetSheet.Cells(GT, TargetCol).Formula = Replace(FML_ST, TEMP_WORD, StRngStr)
        Next SubCnt
    Next MainCnt
    
    'オブジェクト解放
    Set TargetSheet = Nothing
    Set CopyRng = Nothing
    
    'ファイルを閉じる
    TargetBook.Data.Close
    Set TargetBook.Data = Nothing
End Sub

' ----------------------------------------------------------------------
'   関数名：EditHeader（ヘッダー設定処理）
'   引数：  TargetSheet（対象エクセルシート）
'   戻り値：なし
' ----------------------------------------------------------------------
Sub EditHeader(TargetSheet As Worksheet)
    '変数宣言
    Dim Cnt As Long
    Dim TargetCol As Long
    Dim TargetDate As Datelist
    Dim HeaderDate As String
    Dim HeadStr As String
    
    '日付を取得
    With TargetDate
        .BaseDate = Date
        .BaseDateStr = Format(.BaseDate, STR_YMD)
    End With
    
    'タイトルを設定
    With TargetSheet.Range(TITLE_POS)
        .Value = Replace(.Value, TEMP_WORD, SHEET_TITLE & "_" & TargetDate.BaseDateStr)
    End With
    
    'ヘッダーを設定
    For Cnt = 0 To HEAD_CNT
        TargetCol = ITEM_END - (PASTE_COEF * Cnt)
        If (Cnt = 0) Then
            HeaderDate = Format(DateAdd("m", HEAD_CNT, TargetDate.BaseDate), YM)
            HeadStr = HeaderDate & LAST_MO
        ElseIf (Cnt = HEAD_CNT) Then
            HeaderDate = Format(TargetDate.BaseDate, YM)
            HeadStr = HeaderDate & THIS_MO
        Else
            HeaderDate = Format(DateAdd("m", HEAD_CNT - Cnt, TargetDate.BaseDate), YM)
            HeadStr = HeaderDate & "（" & HEAD_CNT - Cnt & MO_UNIT
        End If
        TargetSheet.Cells(HEAD_MO, TargetCol).Value = HeadStr
    Next Cnt
End Sub

' ----------------------------------------------------------------------
'   関数名：EditLine（罫線設定処理）
'   引数：  TargetRng（対象範囲）
'   戻り値：なし
' ----------------------------------------------------------------------
Sub EditLine(TargetRng As Range)
    With TargetRng
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
End Sub

' ----------------------------------------------------------------------
'   関数名：MakeSubTotal（小計行作成処理）
'   引数1： TargetRng（対象範囲）
'   引数2： ProcFlg（列挙型FLGの処理フラグ）
'   戻り値：なし
' ----------------------------------------------------------------------
Sub MakeSubTotal(TargetRng As Range, ProcFlg As Long)
    '変数宣言
    Dim TargetCell As Range
    Dim StCell As Range
    Dim EndCell As Range
    Dim BrRng As Range
    Dim TeamRng As Range
    Dim LastValue As Variant
    
    For Each TargetCell In TargetRng
        If (LastValue <> TargetCell.Value) Then
            '値を検索
            With TargetRng
                Set EndCell = .Find(What:=TargetCell.Value, SearchDirection:=xlPrevious, LookAt:=xlWhole)
                Set StCell = .Find(What:=TargetCell.Value, After:=EndCell, SearchDirection:=xlNext, LookAt:=xlWhole)
            End With
            
            '数式と書式の設定
            Call EditFormulaAndFormat(StCell, EndCell, ProcFlg)
            
            If (ProcFlg = FLG.BR_ST) Then
                '対象チーム範囲を取得
                With ActiveSheet
                    Set TeamRng = .Range(.Cells(StCell.Row, OUTPUT.TEAM), .Cells(EndCell.Row, OUTPUT.TEAM))
                End With
                '再帰処理
                Call MakeSubTotal(TeamRng, FLG.TEAM_ST)
                
                '対象支店範囲を取得
                With ActiveSheet
                    Set BrRng = .Range(.Cells(StCell.Row, OUTPUT.CODE), .Cells(EndCell.Row, OUTPUT.BRANCH))
                End With
                '書式を設定
                With BrRng
                    .Offset(-1).Resize(.Rows.Count + 1).Rows.Group
                    .Offset(-1).Resize(.Rows.Count + 1).Font.ColorIndex = GREY
                    .Offset(-2).Resize(.Rows.Count + 2).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
                End With
            End If
            
            '値を保存
            LastValue = TargetCell.Value
        End If
    Next TargetCell
    
    'オブジェクト解放
    Set TargetRng = Nothing
    Set TargetCell = Nothing
    Set StCell = Nothing
    Set EndCell = Nothing
    Set BrRng = Nothing
    Set TeamRng = Nothing
End Sub

' ----------------------------------------------------------------------
'   関数名：EditFormulaAndFormat（数式と書式の設定処理）
'   引数1： RngSt（範囲始点）
'   引数2： RngEnd（範囲終点）
'   引数3： ProcFlg（列挙型FLGの処理フラグ）
'   戻り値：なし
' ----------------------------------------------------------------------
Sub EditFormulaAndFormat(RngSt As Range, RngEnd As Range, ProcFlg As Long)
    '変数宣言
    Dim MainCnt As Long
    Dim SubCnt As Long
    Dim TargetCol As Long
    Dim ColChar As String
    Dim RowFml As Long
    Dim RowSt As Long
    Dim RowEnd As Long
    Dim BaseRng As Range
    Dim SubtRng As String
    
    '行を挿入
    RngSt.EntireRow.Insert
    '始点と終点の行番号
    RowSt = RngSt.Row
    RowEnd = RngEnd.Row
    '挿入行の行番号（数式入力行）
    RowFml = RowSt - 1
    
    '基準範囲を取得
    With ActiveSheet
        Set BaseRng = .Range(.Cells(RowSt, OUTPUT.CODE), .Cells(RowEnd, OUTPUT.CODE))
    End With
    
    With BaseRng
        If (ProcFlg = FLG.BR_ST) Then
            '挿入行にコピーした値を貼り付け
            .Resize(1, 2).Copy
            .Offset(-1).Resize(1).PasteSpecial Paste:=xlPasteValues
            '書式の設定を初期化
            .Offset(-1).Resize(1, 4).ClearFormats
            '文字を太字に設定
            With .Offset(-1, 2).Resize(1)
                .Value = SUM_BR
                .Font.Bold = True
            End With
            '罫線設定
            Call EditLine(.Offset(-1).Resize(1, 4))
            .Offset(-1, 2).Resize(1, 2).Borders(xlInsideVertical).LineStyle = xlLineStyleNone
        ElseIf (ProcFlg = FLG.TEAM_ST) Then
            '挿入行にコピーした値を貼り付け
            .Resize(1, 3).Copy
            .Offset(-1).Resize(1).PasteSpecial Paste:=xlPasteValues
            '書式の設定を初期化
            .Offset(-1, 2).Resize(1, 2).ClearFormats
             '文字を太字に設定
            With .Offset(-1, 3).Resize(1)
                .Value = SUM_TEAM
                .Font.Bold = True
            End With
            '罫線設定
            Call EditLine(.Offset(-1, 2).Resize(, 2))
            .Offset(-1, 2).Resize(.Rows.Count + 1).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
            With .Offset(, 3).Resize(, 4)
                .Borders(xlInsideHorizontal).LineStyle = xlDash
                .Borders(xlInsideHorizontal).Weight = xlThin
                'グループ化
                .Rows.Group
            End With
            '小計行以外の文字色を設定
            .Offset(, 2).Font.ColorIndex = GREY
        End If
    End With
        
    '挿入行の幅を設定
    With ActiveSheet.Rows(RowFml)
        .EntireRow.AutoFit
        .ClearOutline
    End With

    For MainCnt = 0 To ITEM_CNT
        '対象の列番号取得
        TargetCol = OUTPUT.BOOK + (PASTE_COEF * MainCnt)
        
        '数式を入力
        For SubCnt = 0 To ELEM_CNT
            With ActiveSheet.Cells(RowFml, TargetCol + SubCnt)
                ColChar = Replace(.Address(True, False), "$" & RowFml, "")
                SubtRng = ColChar & RowSt & ":" & ColChar & RowEnd
                .Formula = Replace(FML_ST, TEMP_WORD, SubtRng)
            End With
        Next SubCnt
        
        '基準範囲を取得
        With ActiveSheet
            Set BaseRng = .Range(.Cells(RowSt, TargetCol), .Cells(RowEnd, TargetCol))
        End With
        
        '書式の設定
        With BaseRng
            If (ProcFlg = FLG.BR_ST) Then
                '書式の設定を初期化
                .Offset(-1).Resize(1, 3).ClearFormats
                '罫線設定
                Call EditLine(.Offset(-1).Resize(1, 3))
            ElseIf (ProcFlg = FLG.TEAM_ST) Then
                '罫線設定
                With .Resize(, 3)
                    .Borders(xlInsideHorizontal).LineStyle = xlDash
                    .Borders(xlInsideHorizontal).Weight = xlThin
                End With
            End If
            '文字を太字に設定
            .Offset(-1).Resize(1, 3).Font.Bold = True
        End With
    Next MainCnt
    
    'オブジェクト解放
    Set BaseRng = Nothing
End Sub

' ----------------------------------------------------------------------
'   関数名：EditFormat（書式設定処理）
'   引数：  TargetSheet（対象エクセルシート）
'   戻り値：なし
' ----------------------------------------------------------------------
Sub EditFormat(TargetSheet As Worksheet)
    '変数宣言
    Dim Cnt As Long
    Dim BaseRng As Range
    Dim EditStPos As Long
    
    '基準範囲を取得
    With TargetSheet
        Set BaseRng = .Range(.Cells(BEGIN, OUTPUT.CODE), .Cells(Rows.Count, OUTPUT.CODE).End(xlUp))
    End With
        
    For Cnt = 0 To ITEM_CNT
        EditStPos = PASTE_COEF * (Cnt + 1)
        '書式設定
        With BaseRng.Offset(, EditStPos)
            .Resize(, 2).NumberFormatLocal = COMMA
            .Offset(, 2).NumberFormatLocal = THOU
        End With
        '条件付き書式設定（値が0の場合、文字色を灰色）
        With BaseRng.Offset(, EditStPos).Resize(, 3)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
            .FormatConditions(1).Font.ColorIndex = GREY
        End With
    Next Cnt
    '値を中央揃え
    BaseRng.Resize(, 2).HorizontalAlignment = xlCenter
    'グループ行を折り畳み
    TargetSheet.Outline.ShowLevels RowLevels:=2
    
    'オブジェクト解放
    Set BaseRng = Nothing
End Sub

' ----------------------------------------------------------------------
'   関数名：EndProc（マクロ終了処理）
'   引数：  なし
'   戻り値：なし
' ----------------------------------------------------------------------
Sub EndProc()
    With Application
        '画面更新再開
        .ScreenUpdating = True
    
        '確認メッセージ表示
        .DisplayAlerts = True
    End With
    
    '終了メッセージ
    MsgBox MSG_END, vbInformation
End Sub
