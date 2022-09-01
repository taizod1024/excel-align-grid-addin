Attribute VB_Name = "Module1"
Option Explicit

Sub �O���b�h���ɑ�����_onAction(constrol As IRibbonControl)

    �O���b�h���ɑ�����

End Sub

Sub �Б��ڑ��̃R�l�N�^_onAction(constrol As IRibbonControl)

    �Б��ڑ��̃R�l�N�^

End Sub

Sub �O���b�h���ɑ�����_getEnabled(control As IRibbonControl, ByRef enabled)

    enabled = Not (ActiveWindow Is Nothing)
    
End Sub

Sub �Б��ڑ��̃R�l�N�^_getEnabled(control As IRibbonControl, ByRef enabled)

    enabled = Not (ActiveWindow Is Nothing)

End Sub

Sub �O���b�h���ɑ�����()

    Dim oBefore As Object   ' �Z���N�V�����̕ۑ��p
    Dim lColMax As Long     ' �V�[�g�̍ő��
    Dim lRowMax As Long     ' �V�[�g�̍ő�s��
    Dim dblXMax As Double   ' X���̍ő�l
    Dim dblYMax As Double   ' Y���̍ő�l
    Dim dblWork As Double   ' X���EY���v�Z�p
    Dim lIdxCol As Long     ' ��񋓗p
    Dim lIdxRow As Long     ' �s�񋓗p
    Dim dblXArr() As Double ' X���̔z��
    Dim dblYArr() As Double ' Y���̔z��
    Dim bTgt As Boolean     ' �������̐}�`������Ώۂ��ǂ����̃t���O
    Dim bChg As Boolean     ' �������̐}�`�𑀍삵���ǂ����̃t���O
    Dim lCnt As Long        ' �}�`�̌�
    Dim lShp As Long        ' ����Ώۂ̐}�`�̌�
    Dim lChg As Long        ' �C�������}�`�̌�
    Dim oShp As Shape       ' ����Ώۂ̐}�`
    Dim oShpRng As ShapeRange ' ����Ώۂ̐}�`�S��
    Dim sInfo As String     ' ���b�Z�[�W
    
    ' ���s�O�`�F�b�N
    If TypeName(Selection) <> "Range" Then
        Set oShpRng = Selection.ShapeRange
        ' ���s�m�F
        Select Case MsgBox("���̑���͌��ɖ߂��܂���" + vbCrLf + "�I������Ă���" & oShpRng.Count & "�̐}�`���O���b�h���ɑ����܂����H", _
            vbOKCancel + vbExclamation)
        Case vbOK
        Case vbCancel
            Exit Sub
        End Select
    Else
        ' ���s�m�F
        Select Case MsgBox("���̑���͌��ɖ߂��܂���" + vbCrLf + "���ׂĂ̐}�`���O���b�h���ɑ����܂����H", _
            vbOKCancel + vbExclamation)
        Case vbOK
        Case vbCancel
            Exit Sub
        End Select
        ' �I�����Đ}�`�����邱�Ƃ��m�F
        If ActiveSheet.Shapes.Count = 0 Then
            MsgBox "�}�`�͂���܂���", vbInformation
            Exit Sub
        End If
        ActiveSheet.Shapes.SelectAll
        Set oShpRng = Selection.ShapeRange
    End If
    
    ' ��ʕ`���~
    Application.ScreenUpdating = False
    
    ' �Z���̗񐔁E�s���̎擾
    Set oBefore = Selection
    Cells.Select
    lColMax = Cells.Columns.Count
    lRowMax = Cells.Rows.Count
    If TypeOf oBefore Is Range Then
        oBefore.Select
    Else
        ActiveWindow.VisibleRange(1, 1).Select
    End If
    
    ' ��ʕ`��ĊJ
    Application.ScreenUpdating = True
    
    ' X���̍ő�l�EY���̍ő�l�̎擾
    dblXMax = 0
    dblYMax = 0
    For lCnt = 1 To oShpRng.Count
        Set oShp = oShpRng.Item(lCnt)
        dblWork = oShp.Left + oShp.Width
        If dblXMax < dblWork Then dblXMax = dblWork
        dblWork = oShp.Top + oShp.Height
        If dblYMax < dblWork Then dblYMax = dblWork
    Next
        
    ' ��P�ʂɉ��ʒu�̎擾�E�s�P�ʂɏc�ʒu�̎擾
    lIdxCol = 0
    Do
        lIdxCol = lIdxCol + 1
        ReDim Preserve dblXArr(lIdxCol) As Double
        dblWork = ActiveSheet.Cells(1, lIdxCol).Left
        dblXArr(lIdxCol) = dblWork
    Loop While dblWork < dblXMax And lIdxCol < lColMax
    
    lIdxRow = 0
    Do
        lIdxRow = lIdxRow + 1
        ReDim Preserve dblYArr(lIdxRow) As Double
        dblWork = ActiveSheet.Cells(lIdxRow, 1).Top
        dblYArr(lIdxRow) = dblWork
    Loop While dblWork < dblYMax And lIdxRow < lRowMax
    
    ' ���ׂĂ̐}�`�ɑ΂��ċ߂��O���b�h���Ɋ񂹂�
    lShp = 0
    lChg = 0
    For lCnt = 1 To oShpRng.Count
        
        Set oShp = oShpRng.Item(lCnt)
        
        ' ���b�Z�[�W
        Application.StatusBar = "�S" + CStr(oShpRng.Count) + "�̐}�`��" + CStr(lCnt) + "�Ԗڂ𒲐��� ..."
        DoEvents
        
        bChg = False
        ' ����Ώۂ��ǂ����𔻒f
        If oShp.Connector Then
            ' �R�l�N�^�Ȃ�...
            ' �����Ƃ��q�����Ă��Ȃ�����
            bTgt = _
                Not oShp.ConnectorFormat.BeginConnected And _
                Not oShp.ConnectorFormat.EndConnected
        Else
            ' �R�l�N�^�łȂ��Ȃ�...
            ' �R�����g�ł͂Ȃ�����
            bTgt = _
               oShp.Type <> msoComment
        End If
        ' ����ΏۂȂ�...
        If bTgt Then
            lShp = lShp + 1
            With oShp
            
                ' Left�̒���
                For lIdxCol = 1 To UBound(dblXArr) - 1
                    If dblXArr(lIdxCol) = .Left Then
                        Exit For
                    ElseIf dblXArr(lIdxCol) < .Left And .Left < dblXArr(lIdxCol + 1) Then
                        bChg = True
                        If .Left < (dblXArr(lIdxCol) + dblXArr(lIdxCol + 1)) / 2 Then
                            .LockAspectRatio = False
                            .Left = dblXArr(lIdxCol)
                        Else
                            .LockAspectRatio = False
                            .Left = dblXArr(lIdxCol + 1)
                        End If
                        Exit For
                    End If
                Next
                dblWork = .Left
                
                ' Width�̒���
                For lIdxCol = lIdxCol To UBound(dblXArr) - 1
                    If dblXArr(lIdxCol) = .Left + .Width Then
                        Exit For
                    ElseIf dblXArr(lIdxCol) < .Left + .Width And .Left + .Width < dblXArr(lIdxCol + 1) Then
                        bChg = True
                        If .Left + .Width < (dblXArr(lIdxCol) + dblXArr(lIdxCol + 1)) / 2 Then
                            .LockAspectRatio = False
                            .Width = dblXArr(lIdxCol) - .Left
                        Else
                            .LockAspectRatio = False
                            .Width = dblXArr(lIdxCol + 1) - .Left
                        End If
                        Exit For
                    End If
                Next

                ' Left�̍Ē���
                ' ���̂�Width�𒲐������Ƃ��ɁALeft�������ɂ����ꍇ�����邽��...
                .Left = dblWork

                ' Top�̒���
                For lIdxRow = 1 To UBound(dblYArr) - 1
                    If dblYArr(lIdxRow) = .Top Then
                        Exit For
                    ElseIf dblYArr(lIdxRow) < .Top And .Top < dblYArr(lIdxRow + 1) Then
                        bChg = True
                        If .Top < (dblYArr(lIdxRow) + dblYArr(lIdxRow + 1)) / 2 Then
                            .LockAspectRatio = False
                            .Top = dblYArr(lIdxRow)
                        Else
                            .LockAspectRatio = False
                            .Top = dblYArr(lIdxRow + 1)
                        End If
                        Exit For
                    End If
                Next
                dblWork = .Top
                
                ' Height�̒���
                For lIdxRow = lIdxRow To UBound(dblYArr) - 1
                    If dblYArr(lIdxRow) - .Top = .Height Then
                        Exit For
                    ElseIf dblYArr(lIdxRow) - .Top < .Height And .Height < dblYArr(lIdxRow + 1) - .Top Then
                        bChg = True
                        If .Top + .Height < (dblYArr(lIdxRow) + dblYArr(lIdxRow + 1)) / 2 Then
                            .LockAspectRatio = False
                            .Height = dblYArr(lIdxRow) - .Top
                        Else
                            .LockAspectRatio = False
                            .Height = dblYArr(lIdxRow + 1) - .Top
                        End If
                        Exit For
                    End If
                Next
                
                ' Top�̍Ē���
                ' ���̂�Height�𒲐������Ƃ��ɁALeft�������ɂ����ꍇ�����邽��...
                .Top = dblWork
                
                ' �ύX����Ă�����...
                If bChg Then
                    If lChg = 0 Then
                        ActiveSheet.Cells(lIdxRow, lIdxCol).Activate
                    End If
                    ActiveSheet.Shapes(.Name).Select False
                    sInfo = "��:" & CStr(.Left) & " ��:" & CStr(.Top) & " ��:" & CStr(.Width) & " ����:" & CStr(.Height)
                    lChg = lChg + 1
                End If
            End With
        End If
    Next
    
    ' ���b�Z�[�W
    Application.StatusBar = ""
    
    ' �Ō�̃��b�Z�[�W
    If lShp = 0 Then
        MsgBox "�}�`�͂���܂���", vbInformation
    ElseIf lChg <> 0 Then
        MsgBox CStr(lChg) & "�̐}�`���O���b�h���ɑ����܂���" & vbCrLf & sInfo, vbInformation
    Else
        MsgBox "���ׂĂ̐}�`���O���b�h���ɑ����Ă��܂�", vbInformation
    End If
    
End Sub

Sub �Б��ڑ��̃R�l�N�^()

    Dim shp As Shape    ' shape
    Dim flg As Boolean  ' flag
    
    ' �X���C�h�̐}�`�ꗗ
    For Each shp In ActiveSheet.Shapes
        
        If shp.Connector Then
        
            flg = False
            
            ' �Б��R�l�N�^�̃`�F�b�N
            If shp.ConnectorFormat.BeginConnected And Not shp.ConnectorFormat.EndConnected Then flg = True
            If shp.ConnectorFormat.EndConnected And Not shp.ConnectorFormat.BeginConnected Then flg = True
            
            ' �Б��ڑ��̃R�l�N�^������������I��
            If flg Then
                shp.Select
                Exit Sub
            End If
            
        End If
        
    Next
    
    ' �Б��ڑ��̃R�l�N�^��������Ȃ��������Ƃ�ʒm
    MsgBox "�Б��ڑ��̃R�l�N�^�͂���܂���B", vbInformation
    Exit Sub
    
End Sub

