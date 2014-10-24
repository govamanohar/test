Option Strict Off
Option Explicit On
Friend Class frmDefaulTranMaint
    Inherits System.Windows.Forms.Form

    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Dim Index As Short = Command1.GetIndex(eventSender)

        Dim blnTableKey As Object
        Dim arrData() As Object
        Dim intI As Short
        Dim rsUpdateDftTran As VBtoRecordSet
        Dim sSql As String

        On Error GoTo HANDLER

        If Index = 1 Then
            frmDefaultTranCodes.Tag = "CANCEL"
            Me.Close()
            Exit Sub
        End If

        If Text1(0).Text = "" Then
            MsgBox("You must enter a Code before adding the new Trancode", MsgBoxStyle.Information, "Default Trancode Information")
            Text1(0).Focus()
            Exit Sub
        ElseIf Text1(1).Text = "" Then
            MsgBox("You must enter a Trancode number before adding the new Trancode", MsgBoxStyle.Information, "Default Trancode Information")
            Text1(1).Focus()
            Exit Sub
        End If
        ' ACTIONCODE = '" & SQLChar(mskAction(0).Text) & "'"
        ''Opening rs to update DFTTRAN
        rsUpdateDftTran = New VBtoRecordSet()
        ' sSql = "Select * from DFTTRAN where CODE='" & _Text1_0.Text & "'"
        ' sSql = "SELECT CODE, TRANCODE, DESC1, REQUIRED FROM DFTTRAN WHERE CODE='" & Text1(0).Text & "'"
        sSql = "SELECT CODE, TRANCODE, DESC1, REQUIRED FROM DFTTRAN WHERE 1 = 2"

        rsUpdateDftTran.ActiveConnection.ConnectionString = GetConnectionString()
        rsUpdateDftTran.OpenRecordset(sSql, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        ' Dim sd As String = rsUpdateDftTran.RecordCount
        'rsUpdateDftTran = OpenRecordset(sSql, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly
        'sSql = String.Empty
        
        intI = 0
        ''Nullifying the array
        ReDim arrData(3)
        For intI = LBound(arrData) To UBound(arrData)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            'UPGRADE_WARNING: Couldn't resolve default property of object arrData(intI). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            arrData(intI) = System.DBNull.Value
        Next
        ''Filling array
        'UPGRADE_WARNING: Couldn't resolve default property of object arrData(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arrData(0) = Text1(0).Text
        If cboDftTran.Text = "<No Selection>" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object arrData(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            arrData(1) = ""
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object arrData(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            arrData(1) = Mid(cboDftTran.Text, 1, InStr(1, cboDftTran.Text, "-") - 1)
            If Len(arrData(1)) > 6 Then
                MsgBox("The Tran Code field can have up to 6 characters.  Please modify this field.")
                cboDftTran.Focus()
                rsUpdateDftTran.Close()
                Erase arrData
                Exit Sub
            End If
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object arrData(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arrData(2) = Text1(1).Text
        'UPGRADE_WARNING: Couldn't resolve default property of object arrData(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arrData(3) = IIf(VB6.GetItemString(cboRequired, cboRequired.SelectedIndex) = "Yes", -1, 0)



        If Me.Tag = "NEW" Then
            'If rsUpdateDftTran.RecordCount > 0 Then
            '    Dim codeval As String = Me._Text1_0.Text
            '    MessageBox.Show("Default Trancode" + " '" + codeval + " '" + "already exists. Please modify The code field")
            '    Me._Text1_0.Focus()
            '    Exit Sub
            'End If
            'DBBeginTran()
            DBInsertRec("DFTTRAN", rsUpdateDftTran, arrData)
            'DBCommit()
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object frmDefaultTranCodes.varTableKey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sSql = "CODE = '" & frmDefaultTranCodes.varTableKey & "'"
            ''Updating rs
            'DBBeginTran()
            DBUpdateRec("DFTTRAN", rsUpdateDftTran, arrData, sSql)
            'DBCommit()
        End If


        ''Clearing SQL stmt
        sSql = ""
        ''Cleaning up rs
        rsUpdateDftTran.Close()
        'UPGRADE_NOTE: Object rsUpdateDftTran may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rsUpdateDftTran = Nothing

        ''Clear controls
        For intI = Text1.LBound To Text1.UBound
            Text1(intI).Text = ""
        Next
        'cboDftTran = ""
        'cboRequired = ""
        Me.Close()

        Exit Sub
HANDLER:

        ' If Err = -2147217873 Then
        If GetDAOError(3146, "duplicate key", "UNIQUE CONSTRAINT", "PRIMARY KEY constraint") Then
            If goDAOError.OrStr Then
                DBRollback()
                'UPGRADE_WARNING: Couldn't resolve default property of object arrData(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                MsgBox("Default Trancode '" & arrData(0) & "' already exists.  Please modify the Code field.")
                Text1(0).Focus()
                Exit Sub
            End If
        End If

        Select Case Errors("frmDefaulTranMaint.Command1_Click")
            Case MsgBoxResult.Retry
                Resume
            Case MsgBoxResult.Ignore
                Resume Next
            Case MsgBoxResult.Abort
                'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                Exit Sub
        End Select

    End Sub

    'UPGRADE_WARNING: Form event frmDefaulTranMaint.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmDefaulTranMaint_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        If Me.Tag = "NEW" Then
            '	Me.Text = Me.Text & " Add"
            SelectListIndex("<No Selection>", cboDftTran)
            'Else
            '	Me.Text = Me.Text & " Edit"
        End If
        'SelectListIndex("<No Selection>", cboDftTran)
    End Sub

    Public Sub frmDefaulTranMaint_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Dim rsLoadCombo As ADODB.Recordset
        Dim aTranCodes() As ValidTranCodes
        Dim sSql As String
        Dim intI As Short
        Dim a As Integer
        Dim i As Integer
        If Me.Tag = "NEW" Then
            'Me.Text = Me.Text & " Add"
            Me._Text1_0.Focus()
            ' cboDftTran.Items.Add("<No Selection>")
            ''Adding No/Yes to ComboBox
            GetTranCodes(cboDftTran, "", "", aTranCodes)
            cboRequired.Items.Clear()
            cboRequired.Items.Add("No")
            cboRequired.Items.Add("Yes")
            cboRequired.SelectedIndex = 0
            cboDftTran.Items.Add(LoadStr("<No Selection>"))
            ' cboRequired.SelectedIndex = cboRequired.SelectedValue

        End If
        'If Me.Tag = "EDIT" Then
        '    cboRequired.Items.Add("No")
        '    cboRequired.Items.Add("Yes")
        '    ' GetTranCodes(cboDftTran, "", "", aTranCodes)
        '    'If Me.Tag = "EDIT" Then
        '    '    SelectListIndex("strTranCode", cboDftTran)
        'End If

        'SelectListIndex("strTranCode", cboDftTran)


    End Sub

    Private Sub Text1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.Enter
        Dim Index As Short = Text1.GetIndex(eventSender)
        If Text1(Index).Text <> "" Then
            Text1(Index).SelectionStart = 0
            Text1(Index).SelectionLength = Len(Trim(Text1(Index).Text))
        End If
    End Sub
End Class
