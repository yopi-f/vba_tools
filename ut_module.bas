Attribute VB_Name = "ut"
'高速化
Public Sub Auto_update_stop()
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
End Sub
Public Sub Auto_update()
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub
'最終行取得
Public Function Last_row_get(ByVal ws As Worksheet, ByVal col1 As Long) As Long
    Last_row_get = ws.Cells(ws.Rows.Count, col1).End(xlUp).Row
End Function

Public Sub Preprocess(Hws, LR)
    
    Dim i As Long 'ループ処理に使用
    Dim wbm As Workbook 'マスタcsvファイルを格納する
    Set wbm = Workbooks.Open(ThisWorkbook.Path & "\マスタ.csv") 'このマクロファイルと同じパスのmasterファイルを開く。
    
  With Hws '発注予定商品シート
        .Columns("H:M").Delete Shift:=xlToLeft '以前の計算列削除
        .Cells.FormatConditions.Delete '条件付き書式リセット

        '仕入先への送信形式にするため、発注データを抜き出し
            .Range("H2:H" & LR).FormulaR1C1 = _
                "=IF(COUNTIF(RC[-6],""*)*""),TRIM(MID(RC2,FIND(""】"",RC2)+1,FIND(""("",RC2)-(FIND(""】"",RC2)+1))),TRIM(RIGHT(RC2,LEN(RC2)-FIND(""】"",RC2))))"
            .Range("I2:I" & LR).FormulaR1C1 = "=RC3"
            .Range("J2:J" & LR).FormulaR1C1 = "=IFERROR(MID(RC[-8],SEARCH(""JANで発注"",RC[-8])+7,13)*1,RC[-6])"
            .Range("J2:J" & LR).NumberFormatLocal = 0
            .Range("K2:K" & LR).FormulaR1C1 = "物流倉庫"
            .Range("L2:L" & LR).FormulaR1C1 = "=VLOOKUP(TRIM(RC1),'マスタ.csv'!C4:C6,3,0)"
            .Range("M2:M" & LR).FormulaR1C1 = _
                "=IF(COUNTIF(RC9,""*+*""),(LEFT(RC9,FIND(""+"",RC9)-1)+RIGHT(RC9,LEN(RC9)-FIND(""+"",RC9)))*RC12,RC9*RC12)"
            .Range("H2:M" & LR).Value = .Range("H2:M" & LR).Value
            
   End With
            wbm.Close False 'マスタを閉じる。
End Sub

Public Sub Mail_paste_default(Fws, Hws, ID, LR, Company, Attachment)
 Dim arr As Variant: arr = Array("Excel", "PDF")
 Dim i2 As Long
 Dim LR3 As Long
 Dim Mt As Integer: Mt = 0 'フラグを格納
 Dim cnt As Integer: cnt = 3 'fwsの行数をカウント
 
   With Fws
        For i2 = 2 To LR Step 1
            '指定の仕入先なら処理
            If Hws.Cells(i2, 6) = Company Then
                 .Range("A" & cnt & ":E" & cnt).Value = Hws.Range("H" & i2 & ":L" & i2).Value
                  Hws.Cells(i2, 7) = "発注済"
                    cnt = cnt + 1
            End If
        Next i2
    
         .Range("B1:C1").FormulaR1C1 = ID & "-" & Format(Now(), "yy") & "-" & Format(Now(), "mmdd") & "-1"
        .Range("A2:E" & cnt - 1).Sort key1:=.Range("C2"), order1:=xlAscending, Header:=xlYes '昇順
    
       
    
         '原価とJANが同じなら数量統合
        For i2 = cnt - 1 To 4 Step -1
            If .Cells(i2, 3) = .Cells(i2 - 1, 3) And .Cells(i2, 5) = .Cells(i2 - 1, 5) Then
                .Cells(i2 - 1, 6).FormulaR1C1 = _
                    "=IF(COUNTIF(RC2,""*+*""),LEFT(RC2,FIND(""+"",RC2)-1)+LEFT(R[1]C2,FIND(""+"",R[1]C2)-1)&""+""&RIGHT(RC2,LEN(RC2)-FIND(""+"",RC2))+RIGHT(R[1]C2,LEN(R[1]C2)-FIND(""+"",R[1]C2)),RC2+R[1]C2)"
                .Cells(i2 - 1, 2).Value = .Cells(i2 - 1, 6).Value
                .Rows(i2).Delete
                cnt = cnt - 1
            End If
        Next i2
        'LR3 = Last_row_get(Fws, 1) 'フォーマットシートの1列目最終行を最後に取得
        
   End With
        
        'For i2 = LBound(arr) - 1 To UBound(arr) '配列内にある値と、Attachmentが一致した場合ループ処理中にpublic subを抜ける。
            'If Attachment = arr(i) Then
                'Exit Sub
            'End If
        'Next i
        
         'Call Mail_send(Mt, Company, Fws, LR3) 'フラグ、会社名、フォーマットシート、フォーマットシート最終行を渡す
End Sub


Public Sub Attachment_Excel(Fws, Company)
    Dim LR3 As Long: LR3 = Last_row_get(Fws, 1) 'フォーマットシートの1列目最終行を取得
    Dim NewBook As Workbook
    Dim Tws As Worksheet: Set Tws = ThisWorkbook.Sheets("T様発注用紙")
    Dim Mt As Integer: Mt = 0 'フラグを格納
        If Company = "J様" Then
            Set NewBook = Workbooks.Add
            NewBook.Sheets(1).Range("A1:E" & LR3).Value = Fws.Range("A1:E" & LR3).Value '新規ブックに転記
            NewBook.Sheets(1).Range("A1:E" & LR3).Borders.LineStyle = xlContinuous
            NewBook.Sheets(1).Columns("A:D").AutoFit 'カラムの幅調整
            NewBook.Sheets(1).Columns("C:C").NumberFormatLocal = 0 'JANを数値化
        
            '新しく作成したブックを名前を付けて保存
            NewBook.SaveAs FileName:=ThisWorkbook.Path & "\" & Format(Now, "yyyy.mm.dd") & Company & "発注用紙.xlsx"
            '新しく作成したブックを閉じる
            NewBook.Close False
    
            Mt = 1
        ElseIf Company = "T様" Then
            With Tws
                 'JAN
                     .Range("A20:A" & LR3 + 17).FormulaR1C1 = "=フォーマット!R[-17]C3"
                     '商品名
                     .Range("C20:E" & LR3 + 17).FormulaR1C1 = "=フォーマット!R[-17]C1"
                     '数量
                     .Range("F20:F" & LR3 + 17).FormulaR1C1 = "=フォーマット!R[-17]C2"
                     '単価
                     .Range("G20:G" & LR3 + 17).FormulaR1C1 = "=フォーマット!R[-17]C5"
                     '金額
                     .Range("H20:I" & LR3 + 17).FormulaR1C1 = _
                         "=IF(COUNTIF(RC6,""*+*""),RC7*(LEFT(RC6,FIND(""+"",RC6)-1)+RIGHT(RC6,LEN(RC6)-FIND(""+"",RC6))),RC7*RC6)"
                     Set NewBook = Workbooks.Add
                    
                     'シートを新しいブックへコピーする
                     .Cells.Copy
                     NewBook.Sheets(1).Range("A1").PasteSpecial Paste:=xlPasteAll
                     NewBook.Sheets(1).Range("A20:I200").Value = NewBook.Sheets(1).Range("A20:I200").Value
                     Application.CutCopyMode = False
                     NewBook.Sheets(1).Rows(LR3 + 18 & ":200").Delete Shift:=xlUp
                             '選択セルリセット
                     NewBook.Sheets(1).Range("A1").Select
                             '新しく作成したブックを名前を付けて保存
                     NewBook.SaveAs FileName:=ThisWorkbook.Path & "\" & Format(Now, "yyyy.mm.dd") & Company & "発注用紙.xlsx"
                     '新しく作成したブックを閉じる
                     NewBook.Close False
                     .Range("A20:I200").ClearContents
                     
                     Mt = 1
            End With
        End If
    
        If Mt = 1 Then 'フラグがあればメール作成
               Call Mail_send(Mt, Company, Fws, LR3) 'フラグ、会社名、フォーマットシート、フォーマットシート最終行を渡す
         End If
End Sub

Public Sub Attachment_PDF(Fws, Company)
    Dim LR3 As Long: LR3 = Last_row_get(Fws, 1) 'フォーマットシートの1列目最終行を取得
    Dim NewBook As Workbook
    Dim Kws As Worksheet: Set Kws = ThisWorkbook.Sheets("K様発注用紙")
    Dim Ews As Worksheet: Set Ews = ThisWorkbook.Sheets("E様発注用紙")
    Dim Mt As Integer: Mt = 0 'フラグを格納
    Dim cnt As Integer: cnt = 8
    Dim i As Long
    Dim j As Long
    Dim arr As Variant: arr = Array(3, 16, 19)
   
        If Company = "K様" Then
            With Kws
                   For i = 3 To LR3 'フォーマット行ループ
                      For j = 1 To 3 'フォーマット・K様列ループ
                          If j <> 3 Then
                            .Cells(cnt, arr(j - 1)).Value = Fws.Cells(i, j)
                          Else
                           .Cells(cnt, arr(j - 1)).Value = Fws.Cells(i, 5) 'jが3の時に転記場所をずらす(原価)
                          End If
                      Next j
                            cnt = cnt + 1 'K様の行数を+1
                   Next i
                    'PDF印刷
                    .ExportAsFixedFormat Type:=xlTypePDF, _
                    FileName:=ThisWorkbook.Path & "\" & Format(Now, "yyyy.mm.dd") & Company & "発注用紙", _
                    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
                    
                    '発注済み不要データ削除
                    .Range("C8:V" & LR3 + 5).ClearContents
                    Mt = 2
             End With
        ElseIf Company = "E様" Then
             With Ews
                .Range("A19:A" & LR3 + 16).FormulaR1C1 = "=フォーマット!R[-16]C[2]"
                .Range("C19:E" & LR3 + 16).FormulaR1C1 = "=フォーマット!R[-16]C[-2]"
                .Range("F19:F" & LR3 + 16).FormulaR1C1 = "=フォーマット!R[-16]C[-4]"
                .Range("G19:G" & LR3 + 16).FormulaR1C1 = "=フォーマット!R[-16]C[-2]"
                .Range("H19:I" & LR3 + 16).FormulaR1C1 = "=RC[-2]*RC[-1]"
            
                'PDF印刷
                .ExportAsFixedFormat Type:=xlTypePDF, _
                FileName:=ThisWorkbook.Path & "\" & Format(Now, "yyyy.mm.dd") & Company & "発注用紙", _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
  
                .Range("A19:I" & LR3 + 16).ClearContents
                Mt = 2
              End With
         End If
         
         If Mt = 2 Then 'フラグがあればメール作成
               Call Mail_send(Mt, Company, Fws, LR3) 'フラグ、会社名、フォーマットシート、フォーマットシート最終行を渡す
         End If
End Sub


Public Sub Mail_send(Mt, Company, Fws, LR3)
    
    Dim Mws As Worksheet: Set Mws = ThisWorkbook.Sheets("得意先マスター")
    Dim Mto As String 'to
    Dim Mcc As String 'cc
    Dim Mbcc As String 'bcc
    Dim KENMEI As Variant '件名
    Dim HONBUN As Variant '本文
    Dim SHOMEI As Variant '署名
    Dim ol As Object 'メールオブジェクト
    Dim sel As Object 'メールオブジェクト
    Dim Mail As Object 'メール
        With Mws
             i = 1
             Do
                    i = i + 1
                 If Company = .Cells(i, 1) Then
                        Mto = .Cells(i, 5).Value
                        Mcc = .Cells(i, 6).Value
                        Mbcc = .Cells(i, 7).Value
                        KENMEI = .Cells(i, 8).Value
                        HONBUN = .Cells(i, 9).Value
                        SHOMEI = .Cells(i, 10).Value
                End If
             Loop Until Company = .Cells(i, 1).Value 'companyの値とセルの値が同じになるまでループ
         End With
    
            'メール作成
    Set ol = CreateObject("Outlook.Application")
    Set Mail = ol.CreateItem(0)
    
    Mail.Display '画面を表示
    Mail.To = Mto 'アドレス
    Mail.CC = Mcc
    Mail.BCC = Mbcc
    Mail.Subject = KENMEI '件名
    'Mt1ならExcel、2ならPDFを添付
        If Mt = 1 Then
            Mail.Attachments.Add ThisWorkbook.Path & "\" & Format(Now, "yyyy.mm.dd") & Company & "発注用紙.xlsx"
        End If
        If Mt = 2 Then
            Mail.Attachments.Add ThisWorkbook.Path & "\" & Format(Now, "yyyy.mm.dd") & Company & "発注用紙.pdf"
        End If
    
        Set sel = ol.ActiveInspector.WordEditor.Windows(1).Selection
        sel.TypeText HONBUN & vbCrLf & vbCrLf
        Fws.Range("B1:D" & LR3).HorizontalAlignment = xlCenter
        Fws.Range("A1:E" & LR3).Copy  '新規データ
        sel.Paste
        sel.TypeText vbCrLf & vbCrLf & SHOMEI
        Application.CutCopyMode = False
        Fws.Rows("3:" & LR3).Delete Shift:=xlUp
        
End Sub
    
        
