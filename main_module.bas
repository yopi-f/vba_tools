Attribute VB_Name = "���C��"
Sub �������[�����M()
Attribute �������[�����M.VB_ProcData.VB_Invoke_Func = " \n14"

    Call Auto_update_stop '�����̍�����
    Dim Hws As Worksheet: Set Hws = ThisWorkbook.Sheets("�����\�菤�i") '�V�[�g�i�[
    Dim LR As Long: LR = Last_row_get(Hws, 1) '�����\�菤�i�V�[�g��1��ڍŏI�s���擾
    Dim Tws As Worksheet: Set Tws = ThisWorkbook.Sheets("���Ӑ�}�X�^�[") '�V�[�g�i�[
    Dim LR2 As Long: LR2 = Last_row_get(Tws, 1) '���Ӑ�V�[�g��1��ڍŏI�s���擾
    Dim Fws As Worksheet: Set Fws = ThisWorkbook.Sheets("�t�H�[�}�b�g") '�t�H�[�}�b�g
    Dim LR3 As Long

        
    Dim i As Long '���[�v�����p
    Dim Company As String '���Ӑ�
    Dim ID As Long 'ID
    Dim Ordering_method As String '�������@
    Dim Attachment As String '�Y�t
    Dim alert As VbMsgBoxResult: alert = MsgBox("���s���Ă�낵���ł����H", vbYesNo + vbQuestion, "���s�m�F")
    
    If alert = vbYes Then 'Yes�Ȃ珈�����s
    
      '�����O�O����----------------------------------------------
        Call Preprocess(Hws, LR)
        
      '���C��(����)����--------------------------------------------
        With Tws '�����ɉ����Ĕ������[���̐���

                For i = 2 To LR2
                     Company = .Cells(i, 1).Value
                        If WorksheetFunction.CountIf(Hws.Range("F1:F" & LR), Company) > 0 Then
                                ID = .Cells(i, 2).Value
                                Ordering_method = .Cells(i, 3).Value
                                Attachment = .Cells(i, 4).Value
    
                                '�ʏ�`���̃��[��(Excel�̓Y�torPDF�̓Y�t)
                                    If Ordering_method = "���[��" Then '�������@�̊g�����m�ۂ̂���if�Ŏw��B
                                             Call Mail_paste_default(Fws, Hws, ID, LR, Company, Attachment)
                                        If Attachment = "Excel" Then 'Excel���������̏ꍇ
                                            Call Attachment_Excel(Fws, Company)
                                        ElseIf Attachment = "PDF" Then 'PDF���������̏ꍇ
                                            Call Attachment_PDF(Fws, Company)
                                        Else
                                            Mt = 0
                                            LR3 = Last_row_get(Fws, 1) '�t�H�[�}�b�g�V�[�g��1��ڍŏI�s���擾
                                            Call Mail_send(Mt, Company, Fws, LR3) '���ɂȂɂ��w�肪�Ȃ������ꍇ
                                        End If
                                    End If
                        End If
                 Next i
        End With
                  
      '�㏈��--------------------------------------------
        '�����σf�[�^�̐F�t���ƌv�Z��̍폜
        With Hws.Columns("A:F")
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$G1=""������"""
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
                With .FormatConditions(1).Interior
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                End With
            .FormatConditions(1).StopIfTrue = False
        End With
        Hws.Columns("H:M").Delete Shift:=xlToLeft '�v�Z��폜
        Call Auto_update '����������
        MsgBox "�������܂����B"

    End If '���s�m�F

End Sub

