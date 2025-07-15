Attribute VB_Name = "メイン"
Sub 発注メール送信()
Attribute 発注メール送信.VB_ProcData.VB_Invoke_Func = " \n14"

    Call Auto_update_stop '処理の高速化
    Dim Hws As Worksheet: Set Hws = ThisWorkbook.Sheets("発注予定商品") 'シート格納
    Dim LR As Long: LR = Last_row_get(Hws, 1) '発注予定商品シートの1列目最終行を取得
    Dim Tws As Worksheet: Set Tws = ThisWorkbook.Sheets("得意先マスター") 'シート格納
    Dim LR2 As Long: LR2 = Last_row_get(Tws, 1) '得意先シートの1列目最終行を取得
    Dim Fws As Worksheet: Set Fws = ThisWorkbook.Sheets("フォーマット") 'フォーマット
    Dim LR3 As Long

        
    Dim i As Long 'ループ処理用
    Dim Company As String '得意先
    Dim ID As Long 'ID
    Dim Ordering_method As String '発注方法
    Dim Attachment As String '添付
    Dim alert As VbMsgBoxResult: alert = MsgBox("実行してよろしいですか？", vbYesNo + vbQuestion, "実行確認")
    
    If alert = vbYes Then 'Yesなら処理続行
    
      '発注前前処理----------------------------------------------
        Call Preprocess(Hws, LR)
        
      'メイン(中間)処理--------------------------------------------
        With Tws '条件に応じて発注メールの生成

                For i = 2 To LR2
                     Company = .Cells(i, 1).Value
                        If WorksheetFunction.CountIf(Hws.Range("F1:F" & LR), Company) > 0 Then
                                ID = .Cells(i, 2).Value
                                Ordering_method = .Cells(i, 3).Value
                                Attachment = .Cells(i, 4).Value
    
                                '通常形式のメール(Excelの添付orPDFの添付)
                                    If Ordering_method = "メール" Then '発注方法の拡張性確保のためifで指定。
                                             Call Mail_paste_default(Fws, Hws, ID, LR, Company, Attachment)
                                        If Attachment = "Excel" Then 'Excelだった時の場合
                                            Call Attachment_Excel(Fws, Company)
                                        ElseIf Attachment = "PDF" Then 'PDFだった時の場合
                                            Call Attachment_PDF(Fws, Company)
                                        Else
                                            Mt = 0
                                            LR3 = Last_row_get(Fws, 1) 'フォーマットシートの1列目最終行を取得
                                            Call Mail_send(Mt, Company, Fws, LR3) '特になにも指定がなかった場合
                                        End If
                                    End If
                        End If
                 Next i
        End With
                  
      '後処理--------------------------------------------
        '発注済データの色付けと計算列の削除
        With Hws.Columns("A:F")
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$G1=""発注済"""
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
                With .FormatConditions(1).Interior
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                End With
            .FormatConditions(1).StopIfTrue = False
        End With
        Hws.Columns("H:M").Delete Shift:=xlToLeft '計算列削除
        Call Auto_update '高速化解除
        MsgBox "完了しました。"

    End If '実行確認

End Sub

