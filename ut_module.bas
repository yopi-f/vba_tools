Attribute VB_Name = "ut"
'������
Public Sub Auto_update_stop()
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
End Sub
Public Sub Auto_update()
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub
'�ŏI�s�擾
Public Function Last_row_get(ByVal ws As Worksheet, ByVal col1 As Long) As Long
    Last_row_get = ws.Cells(ws.Rows.Count, col1).End(xlUp).Row
End Function

Public Sub Preprocess(Hws, LR)
    
    Dim i As Long '���[�v�����Ɏg�p
    Dim wbm As Workbook '�}�X�^csv�t�@�C�����i�[����
    Set wbm = Workbooks.Open(ThisWorkbook.Path & "\�}�X�^.csv") '���̃}�N���t�@�C���Ɠ����p�X��master�t�@�C�����J���B
    
  With Hws '�����\�菤�i�V�[�g
        .Columns("H:M").Delete Shift:=xlToLeft '�ȑO�̌v�Z��폜
        .Cells.FormatConditions.Delete '�����t���������Z�b�g

        '�d����ւ̑��M�`���ɂ��邽�߁A�����f�[�^�𔲂��o��
            .Range("H2:H" & LR).FormulaR1C1 = _
                "=IF(COUNTIF(RC[-6],""*)*""),TRIM(MID(RC2,FIND(""�z"",RC2)+1,FIND(""("",RC2)-(FIND(""�z"",RC2)+1))),TRIM(RIGHT(RC2,LEN(RC2)-FIND(""�z"",RC2))))"
            .Range("I2:I" & LR).FormulaR1C1 = "=RC3"
            .Range("J2:J" & LR).FormulaR1C1 = "=IFERROR(MID(RC[-8],SEARCH(""JAN�Ŕ���"",RC[-8])+7,13)*1,RC[-6])"
            .Range("J2:J" & LR).NumberFormatLocal = 0
            .Range("K2:K" & LR).FormulaR1C1 = "�����q��"
            .Range("L2:L" & LR).FormulaR1C1 = "=VLOOKUP(TRIM(RC1),'�}�X�^.csv'!C4:C6,3,0)"
            .Range("M2:M" & LR).FormulaR1C1 = _
                "=IF(COUNTIF(RC9,""*+*""),(LEFT(RC9,FIND(""+"",RC9)-1)+RIGHT(RC9,LEN(RC9)-FIND(""+"",RC9)))*RC12,RC9*RC12)"
            .Range("H2:M" & LR).Value = .Range("H2:M" & LR).Value
            
   End With
            wbm.Close False '�}�X�^�����B
End Sub

Public Sub Mail_paste_default(Fws, Hws, ID, LR, Company, Attachment)
 Dim arr As Variant: arr = Array("Excel", "PDF")
 Dim i2 As Long
 Dim LR3 As Long
 Dim Mt As Integer: Mt = 0 '�t���O���i�[
 Dim cnt As Integer: cnt = 3 'fws�̍s�����J�E���g
 
   With Fws
        For i2 = 2 To LR Step 1
            '�w��̎d����Ȃ珈��
            If Hws.Cells(i2, 6) = Company Then
                 .Range("A" & cnt & ":E" & cnt).Value = Hws.Range("H" & i2 & ":L" & i2).Value
                  Hws.Cells(i2, 7) = "������"
                    cnt = cnt + 1
            End If
        Next i2
    
         .Range("B1:C1").FormulaR1C1 = ID & "-" & Format(Now(), "yy") & "-" & Format(Now(), "mmdd") & "-1"
        .Range("A2:E" & cnt - 1).Sort key1:=.Range("C2"), order1:=xlAscending, Header:=xlYes '����
    
       
    
         '������JAN�������Ȃ琔�ʓ���
        For i2 = cnt - 1 To 4 Step -1
            If .Cells(i2, 3) = .Cells(i2 - 1, 3) And .Cells(i2, 5) = .Cells(i2 - 1, 5) Then
                .Cells(i2 - 1, 6).FormulaR1C1 = _
                    "=IF(COUNTIF(RC2,""*+*""),LEFT(RC2,FIND(""+"",RC2)-1)+LEFT(R[1]C2,FIND(""+"",R[1]C2)-1)&""+""&RIGHT(RC2,LEN(RC2)-FIND(""+"",RC2))+RIGHT(R[1]C2,LEN(R[1]C2)-FIND(""+"",R[1]C2)),RC2+R[1]C2)"
                .Cells(i2 - 1, 2).Value = .Cells(i2 - 1, 6).Value
                .Rows(i2).Delete
                cnt = cnt - 1
            End If
        Next i2
        'LR3 = Last_row_get(Fws, 1) '�t�H�[�}�b�g�V�[�g��1��ڍŏI�s���Ō�Ɏ擾
        
   End With
        
        'For i2 = LBound(arr) - 1 To UBound(arr) '�z����ɂ���l�ƁAAttachment����v�����ꍇ���[�v��������public sub�𔲂���B
            'If Attachment = arr(i) Then
                'Exit Sub
            'End If
        'Next i
        
         'Call Mail_send(Mt, Company, Fws, LR3) '�t���O�A��Ж��A�t�H�[�}�b�g�V�[�g�A�t�H�[�}�b�g�V�[�g�ŏI�s��n��
End Sub


Public Sub Attachment_Excel(Fws, Company)
    Dim LR3 As Long: LR3 = Last_row_get(Fws, 1) '�t�H�[�}�b�g�V�[�g��1��ڍŏI�s���擾
    Dim NewBook As Workbook
    Dim Tws As Worksheet: Set Tws = ThisWorkbook.Sheets("T�l�����p��")
    Dim Mt As Integer: Mt = 0 '�t���O���i�[
        If Company = "J�l" Then
            Set NewBook = Workbooks.Add
            NewBook.Sheets(1).Range("A1:E" & LR3).Value = Fws.Range("A1:E" & LR3).Value '�V�K�u�b�N�ɓ]�L
            NewBook.Sheets(1).Range("A1:E" & LR3).Borders.LineStyle = xlContinuous
            NewBook.Sheets(1).Columns("A:D").AutoFit '�J�����̕�����
            NewBook.Sheets(1).Columns("C:C").NumberFormatLocal = 0 'JAN�𐔒l��
        
            '�V�����쐬�����u�b�N�𖼑O��t���ĕۑ�
            NewBook.SaveAs FileName:=ThisWorkbook.Path & "\" & Format(Now, "yyyy.mm.dd") & Company & "�����p��.xlsx"
            '�V�����쐬�����u�b�N�����
            NewBook.Close False
    
            Mt = 1
        ElseIf Company = "T�l" Then
            With Tws
                 'JAN
                     .Range("A20:A" & LR3 + 17).FormulaR1C1 = "=�t�H�[�}�b�g!R[-17]C3"
                     '���i��
                     .Range("C20:E" & LR3 + 17).FormulaR1C1 = "=�t�H�[�}�b�g!R[-17]C1"
                     '����
                     .Range("F20:F" & LR3 + 17).FormulaR1C1 = "=�t�H�[�}�b�g!R[-17]C2"
                     '�P��
                     .Range("G20:G" & LR3 + 17).FormulaR1C1 = "=�t�H�[�}�b�g!R[-17]C5"
                     '���z
                     .Range("H20:I" & LR3 + 17).FormulaR1C1 = _
                         "=IF(COUNTIF(RC6,""*+*""),RC7*(LEFT(RC6,FIND(""+"",RC6)-1)+RIGHT(RC6,LEN(RC6)-FIND(""+"",RC6))),RC7*RC6)"
                     Set NewBook = Workbooks.Add
                    
                     '�V�[�g��V�����u�b�N�փR�s�[����
                     .Cells.Copy
                     NewBook.Sheets(1).Range("A1").PasteSpecial Paste:=xlPasteAll
                     NewBook.Sheets(1).Range("A20:I200").Value = NewBook.Sheets(1).Range("A20:I200").Value
                     Application.CutCopyMode = False
                     NewBook.Sheets(1).Rows(LR3 + 18 & ":200").Delete Shift:=xlUp
                             '�I���Z�����Z�b�g
                     NewBook.Sheets(1).Range("A1").Select
                             '�V�����쐬�����u�b�N�𖼑O��t���ĕۑ�
                     NewBook.SaveAs FileName:=ThisWorkbook.Path & "\" & Format(Now, "yyyy.mm.dd") & Company & "�����p��.xlsx"
                     '�V�����쐬�����u�b�N�����
                     NewBook.Close False
                     .Range("A20:I200").ClearContents
                     
                     Mt = 1
            End With
        End If
    
        If Mt = 1 Then '�t���O������΃��[���쐬
               Call Mail_send(Mt, Company, Fws, LR3) '�t���O�A��Ж��A�t�H�[�}�b�g�V�[�g�A�t�H�[�}�b�g�V�[�g�ŏI�s��n��
         End If
End Sub

Public Sub Attachment_PDF(Fws, Company)
    Dim LR3 As Long: LR3 = Last_row_get(Fws, 1) '�t�H�[�}�b�g�V�[�g��1��ڍŏI�s���擾
    Dim NewBook As Workbook
    Dim Kws As Worksheet: Set Kws = ThisWorkbook.Sheets("K�l�����p��")
    Dim Ews As Worksheet: Set Ews = ThisWorkbook.Sheets("E�l�����p��")
    Dim Mt As Integer: Mt = 0 '�t���O���i�[
    Dim cnt As Integer: cnt = 8
    Dim i As Long
    Dim j As Long
    Dim arr As Variant: arr = Array(3, 16, 19)
   
        If Company = "K�l" Then
            With Kws
                   For i = 3 To LR3 '�t�H�[�}�b�g�s���[�v
                      For j = 1 To 3 '�t�H�[�}�b�g�EK�l�񃋁[�v
                          If j <> 3 Then
                            .Cells(cnt, arr(j - 1)).Value = Fws.Cells(i, j)
                          Else
                           .Cells(cnt, arr(j - 1)).Value = Fws.Cells(i, 5) 'j��3�̎��ɓ]�L�ꏊ�����炷(����)
                          End If
                      Next j
                            cnt = cnt + 1 'K�l�̍s����+1
                   Next i
                    'PDF���
                    .ExportAsFixedFormat Type:=xlTypePDF, _
                    FileName:=ThisWorkbook.Path & "\" & Format(Now, "yyyy.mm.dd") & Company & "�����p��", _
                    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
                    
                    '�����ςݕs�v�f�[�^�폜
                    .Range("C8:V" & LR3 + 5).ClearContents
                    Mt = 2
             End With
        ElseIf Company = "E�l" Then
             With Ews
                .Range("A19:A" & LR3 + 16).FormulaR1C1 = "=�t�H�[�}�b�g!R[-16]C[2]"
                .Range("C19:E" & LR3 + 16).FormulaR1C1 = "=�t�H�[�}�b�g!R[-16]C[-2]"
                .Range("F19:F" & LR3 + 16).FormulaR1C1 = "=�t�H�[�}�b�g!R[-16]C[-4]"
                .Range("G19:G" & LR3 + 16).FormulaR1C1 = "=�t�H�[�}�b�g!R[-16]C[-2]"
                .Range("H19:I" & LR3 + 16).FormulaR1C1 = "=RC[-2]*RC[-1]"
            
                'PDF���
                .ExportAsFixedFormat Type:=xlTypePDF, _
                FileName:=ThisWorkbook.Path & "\" & Format(Now, "yyyy.mm.dd") & Company & "�����p��", _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
  
                .Range("A19:I" & LR3 + 16).ClearContents
                Mt = 2
              End With
         End If
         
         If Mt = 2 Then '�t���O������΃��[���쐬
               Call Mail_send(Mt, Company, Fws, LR3) '�t���O�A��Ж��A�t�H�[�}�b�g�V�[�g�A�t�H�[�}�b�g�V�[�g�ŏI�s��n��
         End If
End Sub


Public Sub Mail_send(Mt, Company, Fws, LR3)
    
    Dim Mws As Worksheet: Set Mws = ThisWorkbook.Sheets("���Ӑ�}�X�^�[")
    Dim Mto As String 'to
    Dim Mcc As String 'cc
    Dim Mbcc As String 'bcc
    Dim KENMEI As Variant '����
    Dim HONBUN As Variant '�{��
    Dim SHOMEI As Variant '����
    Dim ol As Object '���[���I�u�W�F�N�g
    Dim sel As Object '���[���I�u�W�F�N�g
    Dim Mail As Object '���[��
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
             Loop Until Company = .Cells(i, 1).Value 'company�̒l�ƃZ���̒l�������ɂȂ�܂Ń��[�v
         End With
    
            '���[���쐬
    Set ol = CreateObject("Outlook.Application")
    Set Mail = ol.CreateItem(0)
    
    Mail.Display '��ʂ�\��
    Mail.To = Mto '�A�h���X
    Mail.CC = Mcc
    Mail.BCC = Mbcc
    Mail.Subject = KENMEI '����
    'Mt1�Ȃ�Excel�A2�Ȃ�PDF��Y�t
        If Mt = 1 Then
            Mail.Attachments.Add ThisWorkbook.Path & "\" & Format(Now, "yyyy.mm.dd") & Company & "�����p��.xlsx"
        End If
        If Mt = 2 Then
            Mail.Attachments.Add ThisWorkbook.Path & "\" & Format(Now, "yyyy.mm.dd") & Company & "�����p��.pdf"
        End If
    
        Set sel = ol.ActiveInspector.WordEditor.Windows(1).Selection
        sel.TypeText HONBUN & vbCrLf & vbCrLf
        Fws.Range("B1:D" & LR3).HorizontalAlignment = xlCenter
        Fws.Range("A1:E" & LR3).Copy  '�V�K�f�[�^
        sel.Paste
        sel.TypeText vbCrLf & vbCrLf & SHOMEI
        Application.CutCopyMode = False
        Fws.Rows("3:" & LR3).Delete Shift:=xlUp
        
End Sub
    
        
