Option Strict Off
Imports VB = Microsoft.VisualBasic
Imports System.Data.OleDb



Friend Class frmTblCode
    Inherits System.Windows.Forms.Form

    Dim frmTblPrime 'Please add Form Name
    Dim sTblName As String = String.Empty
    Dim sTypeSel As String = String.Empty
    Dim sOrder As String = String.Empty
    Dim sSQLWhere As String = String.Empty
    Dim sSql As String = String.Empty
    Dim strold As String
    Dim strnew As String
    Dim actionCode As String
    Dim index As Integer
    Const mcs_Type As String = "STS"
    Const mcs_UIType As String = "NOTE"
    Const sTblSrc As String = "DRAPROF"
    Const mcSPCol_Code As Short = 1
    Const mcSPCol_Paid As Short = 3
    Dim aCol3() As String
    Dim dsSource As New VBtoRecordSet()
    Dim bInitialLoad As Object




    Public Shared Function GetColumnIndex(ByVal dataGrid As DataGridView, ByVal ColumnName As String) As Integer
        Dim i As Integer
        Dim ColumnIndex As Integer = -1
        For i = 0 To dataGrid.Columns.Count - 1
            If dataGrid.Columns(i).Name.ToUpper = ColumnName.ToUpper Then
                ColumnIndex = i
                Exit For
            End If
        Next
        Return ColumnIndex
    End Function
    Private Sub AddTblRec()
        On Error GoTo Err_AddTblRec
        Dim Value As String
        Dim aInput() As Object
        Dim ssSrch As New VBtoRecordSet()
        Dim comboval As Integer

        If cboType.SelectedIndex = 0 Then
            sSql = "select top 1 DFTVAL+1  from NOTECODE where CODE='fed' order by DFTVAL desc"
            ssSrch.OpenRecordset(sSql, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        End If

        ''sSql = "select * from notecode"
        'sSql = "select top 1 DFTVAL+1  from NOTECODE where CODE='fed' order by DFTVAL desc"
        'ssSrch.OpenRecordset(sSql, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        aInput(0) = String.Empty
        aInput(1) = String.Empty
        aInput(2) = String.Empty
        aInput(3) = String.Empty

        DBInsertRec(sTblName, ssSrch, aInput)

        Dim asdf As Integer

        Exit Sub

Err_AddTblRec:

        Select Case Errors("frmTblPrime.cmdOptions_Click")
            Case MsgBoxResult.Retry
                Resume
            Case MsgBoxResult.Ignore
                Resume Next
            Case MsgBoxResult.Abort
                Me.Cursor = System.Windows.Forms.Cursors.Default
                Status_End()
                Exit Sub
        End Select
    End Sub
    '    'UPGRADE_WARNING: Event cboType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboType.SelectedIndexChanged
        If Val(cboType.Tag) = Val(CStr(cboType.SelectedIndex)) Then Exit Sub
        cboType.Tag = cboType.SelectedIndex
        SetHourGlass(True)
        If cboType.SelectedIndex = 3 Then
            '  PaidOff.Visible = True
            sTypeSel = "SELECT TYPEID, CODE, DFTVAL, VALLIMIT FROM  " & sTblName
            '
            sSQLWhere = " Where TYPEID = '" & GetType_Renamed() & "'"
            sOrder = " ORDER BY TYPEID, CODE"

        Else
            ' PaidOff.Visible = False
        End If
        sTypeSel = "SELECT CODE, DFTVAL FROM  " & sTblName
        ' sTypeSel = "SELECT TYPEID, CODE, DFTVAL, VALLIMIT FROM  " & sTblName
        sSQLWhere = " Where TYPEID = '" & GetType_Renamed() & "'"
        sOrder = " ORDER BY TYPEID, CODE"


        Dim a As String = "X,T,L,D,W,E,r,S,A"

        LoadSpread()
        If cboType.SelectedIndex = 3 Then
            'create sql query for getting the exclude value.
            '  sSql=""
            For Each row As DataGridViewRow In sprCodeTable.Rows
                If (String.IsNullOrEmpty(row.Cells("Value").Value) = False) Then
                    If (a.Contains(row.Cells("Value").Value)) Then
                        row.Cells.Item("Paidoff").Value = True
                    End If
                End If
                If row.Cells("Value").Value = "X" OrElse row.Cells("Value").Value = "T" Then
                    row.ReadOnly = True
                End If

            Next
        End If
        ' LoadSpread()
        'If sprCodeTable.MaxRows > 0 Then
        '    sprCodeTable.Row = 1
        'End If
        'UPGRADE_WARNING: Couldn't resolve default property of object bInitialLoad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If bInitialLoad Then
            'UPGRADE_WARNING: Couldn't resolve default property of object bInitialLoad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            bInitialLoad = False
        Else
            sprCodeTable.Focus()
        End If
        SetHourGlass(False)
    End Sub


    'Private Sub cboType_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboType.Leave
    '    sprCodeTable.Tag = 1
    'End Sub
    Private Sub ChgTblRec(ByRef sCodeBef As String, ByRef sDftValBef As String, ByRef sCodeAft As String, ByRef sDftValAft As String)
        Dim sWhere As String = String.Empty

        Dim aInput() As Object
        dsSource.MoveLast()
        ReDim aInput(dsSource.ColumnCount - 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object aInput(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aInput(0) = GetType_Renamed()
        'UPGRADE_WARNING: Couldn't resolve default property of object aInput(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aInput(1) = sCodeAft
        'UPGRADE_WARNING: Couldn't resolve default property of object aInput(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aInput(2) = sDftValAft
        '    aInput(3) = GetTypeLen()
        'UPGRADE_WARNING: Couldn't resolve default property of object aInput(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'aInput(3) = InsretCheckedCell(abc)

        sWhere = "TYPEID = '" & GetType_Renamed() & "' and CODE = '" & SQLChar(aCol1(sprCodeTable.SelectedRows.Count)) & "' "
        DBBeginTran()
        DBUpdateRec(sTblName, dsSource, aInput, sWhere, True)
        DBCommit()
    End Sub
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub
    Private Sub cmdClose_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Leave
        ' sprCodeTable.Tag = sprCodeTable.MaxRows
    End Sub
    Private Sub cmdOptions_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOptions.Click
        Dim Index As Short = cmdOptions.GetIndex(eventSender)
        On Error GoTo Err_cmdOptions_Click
        Dim nCurRow As Object
        Select Case Index
            Case 0 'Add
                AddTblRec()

            Case 1 'Delete
                'UPGRADE_WARNING: Couldn't resolve default property of object nCurRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ' nCurRow = sprCodeTable.Row
                'If sprCodeTable.Lock = False Then
                '    InsretCheckedCell(sprCodeTable.Row, True)
                '    DelTblRec()
                'End If
                'Dim DataTable As New DataTable
                'DataTable.Rows.Add(New String() {"a"}, {"b"})
                'sprCodeTable.DataSource = DataTable
                ' Me.sprCodeTable
                ' sprCodeTable.Focus()
            Case Else
        End Select

        Exit Sub

Err_cmdOptions_Click:
        Select Case Errors("frmTblPrime.cmdOptions_Click")
            Case MsgBoxResult.Retry
                Resume
            Case MsgBoxResult.Ignore
                Resume Next
            Case MsgBoxResult.Abort
                Me.Cursor = System.Windows.Forms.Cursors.Default
                Status_End()
                Exit Sub
        End Select
    End Sub
    '    Private Sub DelTblRec()
    '        On Error GoTo Err_DeltblRec
    '        Dim sWhere As String = String.Empty
    '        Dim sPos As String = String.Empty

    '        SetHourGlass(True)
    '        sprCodeTable.Row = sprCodeTable.ActiveRow
    '        sprCodeTable.Col = 1
    '        sWhere = dsSource.Columns(0).ColumnName & " = '" & GetType_Renamed() & "'"
    '        sWhere = sWhere & " And " & dsSource.Columns(1).ColumnName & " = '" & SQLChar((sprCodeTable.Text)) & "'"

    '        DBBeginTran()
    '        DBDeleteRec(sTblName, sWhere)
    '        DBCommit()
    '        LoadSpread()
    '        SetHourGlass(False)

    '        Exit Sub

    'Err_DeltblRec:

    '        Select Case Errors("frmTblPrime.cmdOptions_Click")
    '            Case MsgBoxResult.Retry
    '                Resume
    '            Case MsgBoxResult.Ignore
    '                Resume Next
    '            Case MsgBoxResult.Abort
    '                Me.Cursor = System.Windows.Forms.Cursors.Default
    '                Status_End()
    '                Exit Sub
    '        End Select
    ' End Sub
    'Private Function FindDup(ByRef sVal As String, ByRef sPrev As String, ByRef iCol As Integer) As String
    '    Dim idx As Short
    '    Dim CurRow As Short
    '    Dim CurCol As Short
    '    Dim sNew As String = String.Empty

    '    Dim sMsg As String = String.Empty


    '    CurRow = sprCodeTable.Row
    '    CurCol = sprCodeTable.Col
    '    Select Case iCol
    '        Case 1 ' Code Column
    '            sNew = VB6.Format(LoadNum(sVal), "00")
    '            sMsg = "Code " & sNew & " already exists.  Please try another!"
    '        Case 2 ' Value
    '            sNew = Trim(sVal)
    '            sMsg = "Value " & sNew & " already exists.  Please try another!"
    '    End Select
    '    sprCodeTable.Col = iCol
    '    For idx = 1 To sprCodeTable.MaxRows
    '        If idx <> CurRow Then
    '            sprCodeTable.Row = idx
    '            If sprCodeTable.Text = sNew Then
    '                FindDup = sPrev
    '                MsgBox(sMsg)
    '                Exit Function
    '            End If
    '        End If
    '    Next

    '    sprCodeTable.Row = CurRow
    '    sprCodeTable.Col = CurCol
    '    FindDup = sNew
    'End Function
    Private Sub frmTblCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        LoadKeyForm(KeyCode, Shift)
    End Sub
    Private Sub frmTblCode_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        lblHidden.Text = gsCargo
        gsCargo = ""

        '    RFC 15526 CHANGES
        '    1. TCLMENUS
        '        A.  ADD VENDOR CODE MENU RECORD
        '        B.  CREATE NEW VENDOR MENU ITEM BETWEEN BORROWER AND NOTES FOR REPORTS - 150 INDEX
        '        C.  MOVE VENDOR REPORTS FROM MISCELLANOUS TO UNDER VENDOR MENU     700 TO 150 index  RVEN 100 & 200
        '        D.  ADD TWO NEW VENDOR REPORTS UNDER VENDOR MENU                   150 index RVEN  300 & 400
        '    2.  VENDORCODE table - 4 FIELDS  - TYPEID, CODE, DFTVAL, VALLIMIT
        '    3.  Add fields to VENDOR table - vtype 10, vclass 10, vstatus 10, APPLICATIONDATE, APPROVALDATE, APPROVALLIMIT, MATURITYDATE, REVIEWDATE
        '    4.  Add fields to SECGROUP table - OS_OPT71, OS_OPT72, OS_OPT73, TS_OPT47, TS_OPT48

        SetHourGlass(True)

        CenterSDIForm(Me)

        If lblHidden.Text = "NOTE" Then
            sTblName = "NOTECODE"
            Me.Text = "Note Code Maintenance"
        ElseIf lblHidden.Text = "BORR" Then
            sTblName = "BORRCODE"
            Me.Text = "Borrower Code Maintenance"
        ElseIf lblHidden.Text = "COUNTRY" Then
            sTblName = "COUNTRY"
            Me.Text = "Country Code Maintenance"
        Else '        VEND
            sTblName = "VENDORCODE"
            Me.Text = "Vendor Code Maintenance"
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object bInitialLoad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        bInitialLoad = True
        If lblHidden.Text = "NOTE" Then
            cboType.Items.Add("Federal Code")
            cboType.Items.Add("Loan Class")
            cboType.Items.Add("Loan Grade")
            cboType.Items.Add("Loan Status")
            cboType.Items.Add("Loan Type")
            cboType.Items.Add("Loan Purpose")
            cboType.Items.Add("Collateral Code")
            'RFC -17463 - Lease
            cboType.Items.Add("Lease Status")
            cboType.Items.Add("Tenant Type")
            cboType.Items.Add("Sales Contract Status")
            'RFC - 30902 - Selection List for Release Item Status field
            cboType.Items.Add("Release Status")

        ElseIf lblHidden.Text = "BORR" Then
            cboType.Items.Add("Type")
            cboType.Items.Add("Class")
            cboType.Items.Add("Credit Grade")
            cboType.Items.Add("Status")
            cboType.Items.Add("Payment Code")
            cboType.Items.Add("Stale Days")
            cboType.Items.Add("Miscellaneous")
            cboType.Items.Add("LTOB")
        ElseIf lblHidden.Text = "COUNTRY" Then
            cboType.Items.Add("Country Code")
        Else
            '        ?Class?, ?Status? and ?Type?
            cboType.Items.Add("Type")
            cboType.Items.Add("Class")
            cboType.Items.Add("Status")
        End If
        cboType.SelectedIndex = 0
        ' sprCodeTable.Columns(1).Visible=False
        ' sprCodeTable.Columns(2).Visible = False
        SetHourGlass(False)
    End Sub
    'UPGRADE_NOTE: GetType was upgraded to GetType_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Function GetType_Renamed() As String

        If lblHidden.Text = "NOTE" Then
            Select Case cboType.SelectedIndex
                Case 0
                    GetType_Renamed = "FED"
                Case 1
                    GetType_Renamed = "LCL"
                Case 2
                    GetType_Renamed = "GRD"
                Case 3
                    GetType_Renamed = "STS"
                Case 4
                    GetType_Renamed = "LTY"
                Case 5
                    GetType_Renamed = "LPR"
                Case 6
                    GetType_Renamed = "LCO"
                Case 7
                    'RFC - 17463 - Lease
                    GetType_Renamed = "LST"
                Case 8
                    GetType_Renamed = "TTY"
                Case 9
                    GetType_Renamed = "SCS"
                    'RFC - 30902 - Selection List for Release Item Status field
                Case 10
                    GetType_Renamed = "RLS"
            End Select
        ElseIf lblHidden.Text = "BORR" Then
            Select Case cboType.SelectedIndex
                Case 0
                    GetType_Renamed = "TYP"
                Case 1
                    GetType_Renamed = "CLS"
                Case 2
                    GetType_Renamed = "CRD"
                Case 3
                    GetType_Renamed = "STS"
                Case 4
                    GetType_Renamed = "PYM"
                Case 5
                    GetType_Renamed = "STL"
                Case 6
                    GetType_Renamed = "MSC"
                Case 7
                    GetType_Renamed = "LTO"
            End Select
        ElseIf lblHidden.Text = "COUNTRY" Then
            GetType_Renamed = "COU"
        Else
            Select Case cboType.SelectedIndex
                Case 0
                    GetType_Renamed = "TYP"
                Case 1
                    GetType_Renamed = "CLS"
                Case 2
                    GetType_Renamed = "STS"
            End Select
        End If
        'GetType_Renamed
        ' Return cboType.SelectedIndex
    End Function
    Private Function GetTypeLen() As Short

        If lblHidden.Text = "NOTE" Then
            Select Case cboType.SelectedIndex
                Case 0 ' FED
                    GetTypeLen = 4
                Case 1 ' LCL
                    GetTypeLen = 3
                Case 2 ' GRD
                    GetTypeLen = 3
                Case 3 ' STS
                    GetTypeLen = 1

                Case 4 ' LTY
                    GetTypeLen = 10
                Case 5 ' LPR
                    GetTypeLen = 10
                Case 6 ' LCO
                    GetTypeLen = 10
                Case 7 ' LCO
                    GetTypeLen = 15
                Case 8 ' LCO
                    GetTypeLen = 20
                Case 9 ' LCO
                    GetTypeLen = 15
                Case 10 ' RLS
                    GetTypeLen = 1
            End Select
        ElseIf lblHidden.Text = "BORR" Then
            Select Case cboType.SelectedIndex
                Case 0 'TYP
                    GetTypeLen = 10
                Case 1 'CLS
                    GetTypeLen = 10
                Case 2 'CRD
                    GetTypeLen = 10
                Case 3 'STS
                    GetTypeLen = 10
                Case 4 'PYM
                    GetTypeLen = 10
                Case 5 'STL
                    GetTypeLen = 10
                Case 6 'MSC
                    GetTypeLen = 10
                Case 7 'LTO
                    GetTypeLen = 10
            End Select
        ElseIf lblHidden.Text = "COUNTRY" Then
            GetTypeLen = 3
        Else
            Select Case cboType.SelectedIndex
                Case 0 'TYP
                    GetTypeLen = 10
                Case 1 'CLS
                    GetTypeLen = 10
                Case 2 'STS
                    GetTypeLen = 10
            End Select
        End If

    End Function
    'Private Sub LoadCell(ByRef Row As Integer, ByRef Col As Short, ByRef Value As String)
    '    sprCodeTable.Row = Row
    '    sprCodeTable.Col = Col
    '    sprCodeTable.Text = Value
    '    If mcs_Type = GetType_Renamed() And mcs_UIType = lblHidden.Text And Col = 3 Then
    '        If Value <> "1" Then
    '            sprCodeTable.Row = Row
    '            sprCodeTable.Value = "1"
    '        Else
    '            sprCodeTable.Row = Row
    '            sprCodeTable.Value = "0"
    '        End If
    '    End If
    '    If mcs_Type = GetType_Renamed() And mcs_UIType = lblHidden.Text And Col = 2 Then
    '        If Value = "X" Or Value = "T" Then
    '            LockSpread(True, Row)
    '        Else
    '            chkExclude(Row, Col)
    '        End If

    '    End If

    'End Sub
    Private Sub LoadSpread()
        Dim intI As Integer
        Dim nRow As Integer
        Dim nCurRow As Short
        Dim dsSource As New VBtoRecordSet()
        '	'    Set dsSource = New ADODB.Recordset
        '	'    dsSource.CursorLocation = adUseClient
        sSql = " Select * from  Notecode "
        ' sSql = sSql & "ORDER BY ACTIONCODE"
        'dsSource.ActiveConnection.ConnectionString = Helper.GetConnectionString()
        dsSource.OpenRecordset(sTypeSel & sSQLWhere & sOrder, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        'New 


        'New 
        sprCodeTable.DataSource = dsSource.Table()

        'dsSource.Column(0).ColumnName = "TYPEID"
        'dsSource.Column(0).ColumnMapping = MappingType.Hidden
        Do Until dsSource.EOF
            intI = sprCodeTable.Rows.Count
            '  Dim col As Integer = GetColumnIndex(sprCodeTable, "CODE")
            ' Dim col1 As Integer = GetColumnIndex(sprCodeTable, "DFTVAL")
            dsSource.MoveNext()

            intI = intI + 1
        Loop
        dsSource.MoveFirst()

        'Dim nRow As Integer

        ''    Dim nCurRow%

        ' dsSource.ActiveConnection.ConnectionString = Helper.GetConnectionString()
        'dsSource.OpenRecordset(sTypeSel & sSQLWhere & sOrder, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)


        '' First Dispose grid to its Original State

        ''SetSpread((True))
        ''If mcs_UIType = lblHidden.Text Then
        ''    SetSpread((False))
        ''End If

        'If dsSource.EOF Then
        '    ReDim aCol1(0)
        '    ReDim aCol2(0)
        'Else
        '    'UPGRADE_WARNING: Lower bound of array aCol1 was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        '    ReDim aCol1(dsSource.RecordCount)
        '    'UPGRADE_WARNING: Lower bound of array aCol2 was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        '    ReDim aCol2(dsSource.RecordCount)
        '    'UPGRADE_WARNING: Lower bound of array aCol3 was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        '    ReDim aCol3(dsSource.RecordCount)
        '    nRow = 1
        '    Do Until dsSource.EOF
        '        aCol1(nRow) = dsSource.Fields("CODE")
        '        aCol2(nRow) = dsSource.Fields("DFTVAL")

        '        LoadCell(nRow, 1, aCol1(nRow))
        '        LoadCell(nRow, 2, aCol2(nRow))
        '        If mcs_Type = GetType_Renamed() And mcs_UIType = lblHidden.Text Then
        '            aCol3(nRow) = dsSource.Fields("VALLIMIT")
        '            LoadCell(nRow, 3, aCol3(nRow))
        '        End If

        '        nRow = nRow + 1
        '        dsSource.MoveNext()
        '    Loop
        'End If
        ' If mcs_Type = GetType_Renamed() And mcs_UIType = lblHidden.Text Then CheckSpread()
    End Sub

    Private Sub frmTblCode_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        'UPGRADE_NOTE: Object frmTblCode may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '		Me = Nothing

    End Sub


    'Private Sub sprCodeTable_Advance(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_AdvanceEvent) Handles sprCodeTable.Advance ' * VB4.0 *
    '    Select Case eventArgs.AdvanceNext
    '        Case True
    '            cmdClose.Focus()
    '        Case False
    '            cboType.Focus()
    '    End Select

    'End Sub
    ' Private Sub sprCodeTable_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles sprCodeTable.ClickEvent ' * VB4.0 *

    'If sprCodeTable.ActiveRow < 1 Then
    '    Exit Sub
    'End If

    'sprCodeTable.Row = sprCodeTable.ActiveRow
    'sprCodeTable.Col = sprCodeTable.ActiveCol
    '  End Sub

    'Private Sub sprCodeTable_EditModeEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_EditModeEvent) Handles sprCodeTable.EditModeEvent ' * VB4.0 *
    '    Dim NewValue As String = String.Empty

    '    Dim CurValue As String = String.Empty

    '    Dim strSrch As String = String.Empty

    '    Dim clValue As Short
    '    Select Case VB6.Format(eventArgs.Mode) & VB6.Format(eventArgs.ChangeMade)
    '        Case "0False" ' Edit Off No Change
    '            ' Don't Update Anything
    '        Case "0True" ' Edit Off Change Made
    '            ' Update Current Record
    '            SetHourGlass(True)

    '            sprCodeTable.Row = Row
    '            sprCodeTable.Col = Col
    '            CurValue = sprCodeTable.Text

    '            '.........Add

    'ed Validation during RFC 29728
    '            If lblHidden.Text = "NOTE" And Trim(cboType.Text) = "Loan Status" Then
    '                '                chkExclude sprCodeTable.Row
    '                Xstr = Fetch_Exclude
    '                strSrch = String.Empty
    '                strSrch = "X,T" & XstrTbl
    '                sprCodeTable.Row = Row
    '                sprCodeTable.Col = 2
    '                CurValue = sprCodeTable.Text
    '                sprCodeTable.Col = 3
    '                clValue = CShort(sprCodeTable.Value)
    '                If InStr(Trim(strSrch), Trim(UCase(CurValue))) <> 0 And Trim(UCase(CurValue)) <> "" And clValue <> 0 Then
    '                    MsgBox("Value " & CurValue & " already exists.  Please try another!")
    '                    sprCodeTable.Col = Col
    '                    sprCodeTable.Row = Row
    '                    sprCodeTable.Text = String.Empty
    '                    CurValue = String.Empty
    '                    strSrch = String.Empty
    '                    Cursor = System.Windows.Forms.Cursors.Default
    '                    sprCodeTable.Focus()
    '                    Exit Sub
    '                End If
    '            End If
    '            sprCodeTable.Row = Row
    '            sprCodeTable.Col = Col
    '            CurValue = sprCodeTable.Text

    '            Select Case Col
    '                Case 1
    '                    NewValue = FindDup(CurValue, aCol1(Row), Col)
    '                    If NewValue <> aCol1(Row) Then
    '                        ChgTblRec(aCol1(Row), aCol2(Row), NewValue, aCol2(Row))
    '                        aCol1(Row) = NewValue
    '                    End If
    '                    LoadSpread() ' Reload Here For Sort Of Code Values
    '                Case 2
    '                    NewValue = FindDup(CurValue, aCol2(Row), Col)
    '                    If NewValue <> aCol2(Row) Then
    '                        ChgTblRec(aCol1(Row), aCol2(Row), aCol1(Row), NewValue)
    '                        aCol2(Row) = NewValue
    '                    End If
    '                    sprCodeTable.Col = Col
    '                    sprCodeTable.Row = Row
    '                    sprCodeTable.Text = NewValue
    '                    CurValue = String.Empty
    '                    Cursor = System.Windows.Forms.Cursors.Default
    '                    sprCodeTable.Focus()
    '                    Exit Sub

    '                Case 3

    '                    ChgTblRec(aCol1(Row), aCol2(Row), aCol1(Row), aCol2(Row))

    '            End Select

    '            SetHourGlass(False)
    '        Case "1False" ' Edit Mode On, Always 0

    '    End Select

    'End Sub
    'Private Sub sprCodeTable_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sprCodeTable.Enter
    '    Dim Row As Integer
    '    Dim Col As Integer
    '    sprCodeTable.Row = -1
    '    sprCodeTable.Col = -1
    '    If CheckWin7Version = False Then
    '        sprCodeTable.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000009)
    '        sprCodeTable.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000006)
    '    Else
    '        sprCodeTable.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
    '        sprCodeTable.ForeColor = System.Drawing.ColorTranslator.FromOle(&H303030)
    '    End If
    '    If CDbl(sprCodeTable.Tag) = 1 Then
    '        Row = 1
    '        Col = 1
    '    Else
    '        Row = sprCodeTable.MaxRows
    '        Col = 2
    '    End If
    '    sprCodeTable.Row = Row
    '    sprCodeTable.Col = Col
    '    sprCodeTable.Action = FPSpreadADO.ActionConstants.ActionActiveCell
    '    sprCodeTable_LeaveCell(sprCodeTable, New AxFPSpreadADO._DSpreadEvents_LeaveCellEvent(-1, -1, Col, Row, False))
    'End Sub
    ' Private Sub sprCodeTable_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles sprCodeTable.LeaveCell ' * VB4.0 *

    'If CheckWin7Version = True Then 'Execute Only for Vista or Win7

    '    SpreadEvent(Col, Row, eventArgs.NewCol, eventArgs.NewRow, eventArgs.Cancel, sprCodeTable, CStr(&H80000004))
    '    Exit Sub
    'End If

    'If Row <> -1 And Col <> -1 Then
    '    sprCodeTable.Row = Row
    '    sprCodeTable.Col = -1
    '    sprCodeTable.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000009)
    '    sprCodeTable.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000006)
    'End If
    'If eventArgs.NewRow <> -1 And eventArgs.NewCol <> -1 Then
    '    sprCodeTable.Row = eventArgs.NewRow
    '    sprCodeTable.Col = -1
    '    sprCodeTable.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
    '    sprCodeTable.Col = eventArgs.NewCol
    '    sprCodeTable.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000006)
    '    sprCodeTable.ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000009)
    'Else
    '    sprCodeTable.Row = Row
    '    sprCodeTable.Col = -1
    '    sprCodeTable.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
    'End If
    'End Sub
    'Private Sub SetSpread(ByVal blnDispose As Boolean)

    '    If blnDispose = False Then
    '        If mcs_Type <> GetType_Renamed() Then
    '            sprCodeTable.MaxRows = 0


    '            sprCodeTable.MaxRows = dsSource.RecordCount
    '            sprCodeTable.Row = 0
    '            sprCodeTable.Col = 1
    '            sprCodeTable.Text = "Code"
    '            sprCodeTable.TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
    '            sprCodeTable.Row = -1
    '            sprCodeTable.CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
    '            sprCodeTable.TypeEditMultiLine = False
    '            sprCodeTable.TypeEditLen = 2
    '            sprCodeTable.set_ColWidth(1, 800) ' 7

    '            sprCodeTable.Row = 0
    '            sprCodeTable.Col = 2
    '            sprCodeTable.Text = "Value"
    '            sprCodeTable.TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
    '            sprCodeTable.Row = -1
    '            sprCodeTable.CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
    '            sprCodeTable.TypeEditCharSet = FPSpreadADO.TypeEditCharSetConstants.TypeEditCharSetASCII
    '            sprCodeTable.TypeEditMultiLine = False
    '            sprCodeTable.TypeEditLen = GetTypeLen()
    '            sprCodeTable.set_ColWidth(2, 1800) ' 15
    '            'RFC - 17463 - Lease
    '            If InStr(",LST,TTY,SCS,", "," & GetType_Renamed() & ",") > 0 Then
    '                sprCodeTable.Row = 0
    '                sprCodeTable.Col = 2
    '                sprCodeTable.Text = "Value"
    '                sprCodeTable.TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
    '                sprCodeTable.Row = -1
    '                sprCodeTable.CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
    '                sprCodeTable.TypeEditCharSet = FPSpreadADO.TypeEditCharSetConstants.TypeEditCharSetAlphanumeric
    '                sprCodeTable.TypeEditMultiLine = False
    '                sprCodeTable.TypeEditLen = GetTypeLen()
    '                sprCodeTable.set_ColWidth(2, 1800) ' 15
    '                'SetMaxWidth
    '            End If

    '        Else
    '            sprCodeTable.MaxRows = 0
    '            sprCodeTable.MaxCols = 3
    '            sprCodeTable.MaxRows = dsSource.RecordCount
    '            sprCodeTable.Row = 0
    '            sprCodeTable.Col = 1
    '            sprCodeTable.Text = "Code" & Space(5) : sprCodeTable.TypeTextWordWrap = False
    '            sprCodeTable.TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
    '            sprCodeTable.Row = -1
    '            sprCodeTable.CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
    '            sprCodeTable.TypeEditMultiLine = False
    '            sprCodeTable.TypeEditLen = 2


    '            sprCodeTable.Row = 0
    '            sprCodeTable.Col = 2
    '            sprCodeTable.Text = "Value" & Space(8) : sprCodeTable.TypeTextWordWrap = False
    '            sprCodeTable.TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
    '            sprCodeTable.Row = -1
    '            sprCodeTable.CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
    '            sprCodeTable.TypeEditCharSet = FPSpreadADO.TypeEditCharSetConstants.TypeEditCharSetASCII
    '            sprCodeTable.TypeEditMultiLine = False
    '            sprCodeTable.TypeEditLen = GetTypeLen()


    '            sprCodeTable.Row = 0

    '            sprCodeTable.Col = 3
    '            sprCodeTable.Text = "Paid-Off" & Space(4) : sprCodeTable.TypeTextWordWrap = False
    '            sprCodeTable.TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
    '            sprCodeTable.Row = -1
    '            sprCodeTable.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
    '            sprCodeTable.TypeEditMultiLine = False
    '            If dsSource.RecordCount <> 0 Then
    '                sprCodeTable.TypeEditLen = GetTypeLen()
    '            End If
    '            SetMaxWidth()

    '        End If


    '    Else
    '        sprCodeTable.MaxRows = 0
    '        sprCodeTable.MaxCols = 2
    '        sprCodeTable.MaxRows = dsSource.RecordCount
    '        sprCodeTable.Row = 0
    '        sprCodeTable.Col = 1
    '        sprCodeTable.Text = "Code"
    '        sprCodeTable.TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
    '        sprCodeTable.Row = -1
    '        sprCodeTable.CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
    '        sprCodeTable.TypeEditCharSet = FPSpreadADO.TypeEditCharSetConstants.TypeEditCharSetASCII
    '        sprCodeTable.TypeEditMultiLine = False
    '        sprCodeTable.TypeEditLen = 2
    '        sprCodeTable.set_ColWidth(1, 800) ' 7

    '        sprCodeTable.Row = 0
    '        sprCodeTable.Col = 2
    '        sprCodeTable.Text = "Value"
    '        sprCodeTable.TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
    '        sprCodeTable.Row = -1
    '        sprCodeTable.CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
    '        sprCodeTable.TypeEditCharSet = FPSpreadADO.TypeEditCharSetConstants.TypeEditCharSetASCII
    '        sprCodeTable.TypeEditMultiLine = False
    '        sprCodeTable.TypeEditLen = GetTypeLen()
    '        sprCodeTable.set_ColWidth(2, 1800) ' 15


    '    End If

    '    sprCodeTable.GridShowHoriz = True
    '    sprCodeTable.GridShowVert = True
    '    sprCodeTable.EditModeReplace = True

    'End Sub
    'Sub SetMaxWidth()

    '    Dim lngCol As Integer
    '    Dim lngHeader As Integer

    '    With sprCodeTable

    '        For lngCol = 1 To .MaxCols
    '            .Row = 0 : .Col = lngCol
    '            lngHeader = .MaxTextCellWidth

    '            If lngHeader > .get_MaxTextColWidth(lngCol) Then
    '                .set_ColWidth(lngCol, lngHeader)
    '            Else
    '                .set_ColWidth(lngCol, .get_MaxTextColWidth(lngCol))
    '            End If

    '        Next

    '    End With

    'End Sub
    'Private Sub CheckSpread()

    '    Dim strSQL As String = String.Empty

    '    Dim NewValue As String = String.Empty

    '    Dim I As Short
    '    Dim J As Short
    '    Dim Xstatus As New VBtoRecordSet()
    '    Dim aInput() As Object
    '    Dim aColVal1(1) As String
    '    Dim aColVal2(1) As String
    '    Dim where As String = String.Empty

    '    'UPGRADE_WARNING: Lower bound of array aCol3 was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    '    ReDim Preserve aCol3(UBound(aCol3) + 1)

    '    strSQL = String.Empty
    '    strSQL = "Select * from NOTECODE where "
    '    strSQL = strSQL & "TYPEID = 'STS' and DFTVAL  IN ('X'" & ",'T')"

    '    Xstatus = New VBtoRecordSet()
    '    Xstatus.ActiveConnection.ConnectionString = Helper.GetConnectionString()
    '    Xstatus.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

    '    If Xstatus.EOF Then


    '        If aCol2(sprCodeTable.Row) = String.Empty Then
    '            aColVal1(I) = aCol1(sprCodeTable.Row)
    '            For I = 0 To 1
    '                Select Case I
    '                    Case 0
    '                        aColVal2(I) = "X"
    '                        aColVal1(I) = VB6.Format(LoadNum(aCol1(UBound(aCol1))) + 1, "00")
    '                        ReDim aInput(dsSource.ColumnCount - 1)
    '                        'UPGRADE_WARNING: Couldn't resolve default property of object aInput(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                        aInput(0) = GetType_Renamed()
    '                        'UPGRADE_WARNING: Couldn't resolve default property of object aInput(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                        aInput(1) = aColVal1(I)
    '                        'UPGRADE_WARNING: Couldn't resolve default property of object aInput(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                        aInput(2) = aColVal2(I)
    '                        'UPGRADE_WARNING: Couldn't resolve default property of object aInput(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                        aInput(3) = 2
    '                        DBBeginTran()
    '                        DBInsertRec(sTblName, dsSource, aInput)
    '                        DBCommit()
    '                    Case 1
    '                        aColVal2(I) = "T"
    '                        sprCodeTable.MaxRows = sprCodeTable.Row + 1
    '                        sprCodeTable.Row = sprCodeTable.Row + 1
    '                        aCol3(UBound(aCol3)) = VB6.Format(LoadNum(aColVal1(I - 1)) + 1, "00")

    '                        aColVal1(I) = aCol3(UBound(aCol3))

    '                        ReDim aInput(dsSource.ColumnCount - 1)
    '                        'UPGRADE_WARNING: Couldn't resolve default property of object aInput(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                        aInput(0) = GetType_Renamed()
    '                        'UPGRADE_WARNING: Couldn't resolve default property of object aInput(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                        aInput(1) = aColVal1(I)
    '                        'UPGRADE_WARNING: Couldn't resolve default property of object aInput(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                        aInput(2) = aColVal2(I)
    '                        'UPGRADE_WARNING: Couldn't resolve default property of object aInput(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                        aInput(3) = 2

    '                        DBBeginTran()
    '                        DBInsertRec(sTblName, dsSource, aInput)
    '                        DBCommit()

    '                End Select

    '            Next I
    '        Else

    '            sprCodeTable.MaxRows = sprCodeTable.Row + 1
    '            sprCodeTable.Row = sprCodeTable.Row + 1
    '            aCol3(UBound(aCol3)) = VB6.Format(LoadNum(aCol1(UBound(aCol1))) + 1, "00")
    '            For I = 0 To 1
    '                aColVal1(I) = aCol3(UBound(aCol3))

    '                Select Case I
    '                    Case 0
    '                        aColVal2(I) = "X"
    '                    Case 1
    '                        aColVal2(I) = "T"
    '                End Select
    '                ReDim aInput(dsSource.ColumnCount - 1)
    '                'UPGRADE_WARNING: Couldn't resolve default property of object aInput(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                aInput(0) = GetType_Renamed()
    '                'UPGRADE_WARNING: Couldn't resolve default property of object aInput(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                aInput(1) = aColVal1(I)
    '                'UPGRADE_WARNING: Couldn't resolve default property of object aInput(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                aInput(2) = aColVal2(I)
    '                'UPGRADE_WARNING: Couldn't resolve default property of object aInput(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                aInput(3) = 2

    '                DBBeginTran()
    '                DBInsertRec(sTblName, dsSource, aInput)
    '                DBCommit()
    '                aCol3(UBound(aCol3)) = VB6.Format(LoadNum(aCol3(UBound(aCol3))) + 1, "00")
    '            Next I
    '        End If
    '    Else

    '        If Xstatus.RecordCount = 2 Then
    '            For I = 0 To Xstatus.RecordCount - 1
    '                '                                'If X or T exist Update the ValLimit field with 2
    '                '                                'To convert for PayOff's
    '                ReDim aInput(Xstatus.ColumnCount - 1)
    '                'UPGRADE_WARNING: Couldn't resolve default property of object aInput(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                aInput(0) = Xstatus.Fields("CODE")
    '                'UPGRADE_WARNING: Couldn't resolve default property of object aInput(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                aInput(1) = Xstatus.Fields("DFTVAL")
    '                'UPGRADE_WARNING: Couldn't resolve default property of object aInput(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                aInput(2) = Xstatus.Fields("TYPEID")
    '                If LoadNum(Xstatus.Fields("VALLIMIT")) <> 2 Then
    '                    'UPGRADE_WARNING: Couldn't resolve default property of object aInput(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                    aInput(3) = 2
    '                    where = String.Empty
    '                    where = "TYPEID = '" & GetType_Renamed() & "' and CODE = '" & SQLChar(Xstatus.Fields("CODE")) & "' "
    '                    DBBeginTran()
    '                    DBUpdateRec(sTblName, Xstatus, aInput, where, True)
    '                    DBCommit()
    '                End If

    '                If Xstatus.EOF = False Then Xstatus.MoveNext()
    '            Next I

    '        Else
    '            For J = 0 To Xstatus.RecordCount - 1

    '                where = String.Empty
    '                where = "TYPEID = '" & GetType_Renamed() & "' and CODE = '" & SQLChar(Xstatus.Fields("CODE")) & "' "

    '                DBBeginTran()
    '                DBDeleteRec(sTblName, where)
    '                DBCommit()
    '                If Xstatus.EOF = False Then Xstatus.MoveNext()
    '            Next J
    '            'Insert X and T Records for Payoff

    '            sprCodeTable.MaxRows = sprCodeTable.Row + 1
    '            sprCodeTable.Row = sprCodeTable.Row + 1
    '            aCol3(UBound(aCol3)) = VB6.Format(LoadNum(aCol1(UBound(aCol1))) + 1, "00")
    '            For I = 0 To 1
    '                aColVal1(I) = aCol3(UBound(aCol3))

    '                Select Case I
    '                    Case 0
    '                        aColVal2(I) = "X"
    '                    Case 1
    '                        aColVal2(I) = "T"
    '                End Select
    '                ReDim aInput(dsSource.ColumnCount - 1)
    '                'UPGRADE_WARNING: Couldn't resolve default property of object aInput(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                aInput(0) = GetType_Renamed()
    '                'UPGRADE_WARNING: Couldn't resolve default property of object aInput(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                aInput(1) = aColVal1(I)
    '                'UPGRADE_WARNING: Couldn't resolve default property of object aInput(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                aInput(2) = aColVal2(I)
    '                'UPGRADE_WARNING: Couldn't resolve default property of object aInput(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                aInput(3) = 2

    '                DBBeginTran()
    '                DBInsertRec(sTblName, dsSource, aInput)
    '                DBCommit()
    '                aCol3(UBound(aCol3)) = VB6.Format(LoadNum(aCol3(UBound(aCol3))) + 1, "00")
    '            Next I



    '        End If

    '    End If
    '    'UPGRADE_NOTE: Object Xstatus may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '    Xstatus = Nothing

    'End Sub
    'Private Sub UpdateRECX(ByRef sDefT As String, ByRef sDftValBef As String, ByRef sCodeVal As String, ByRef sDftValAft As String)
    '    Dim sWhere As String = String.Empty

    '    Dim aInput() As Object
    '    dsSource.MoveLast()
    '    ReDim aInput(dsSource.ColumnCount - 1)
    '    'UPGRADE_WARNING: Couldn't resolve default property of object aInput(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '    aInput(0) = GetType_Renamed()
    '    'UPGRADE_WARNING: Couldn't resolve default property of object aInput(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '    aInput(1) = sCodeVal
    '    'UPGRADE_WARNING: Couldn't resolve default property of object aInput(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '    aInput(2) = sDftValAft
    '    'UPGRADE_WARNING: Couldn't resolve default property of object aInput(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '    aInput(3) = 2

    '    sWhere = "TYPEID = '" & GetType_Renamed() & "' and CODE = '" & SQLChar(aCol1(sprCodeTable.Row)) & "'"
    '    DBBeginTran()
    '    DBUpdateRec(sTblName, dsSource, aInput, sWhere)
    '    DBCommit()
    ''End Sub
    'Private Sub LockSpread(ByRef blnSprLock As Boolean, Optional ByRef rRow As Integer = 0)

    '    If blnSprLock = True Then


    '        sprCodeTable.Columns(3) = mcSPCol_Code : sprCodeTable.Col2 = mcSPCol_Paid
    '        sprCodeTable.Row = rRow : sprCodeTable.Row2 = rRow
    '        s()
    '        sprCodeTable.Lock = True
    '        '        :      sprSpread.BackColor = Me.BackColor
    '        sprCodeTable.BlockMode = False

    '        sprCodeTable.SelectBlockOptions = SS_SELBLOCKOPT_ROWS
    '        sprCodeTable.UserResize = FPSpreadADO.UserResizeConstants.UserResizeColumns
    '        ''                sprCodeTable.MaxRows = sprCodeTable.DataRowCnt
    '    Else
    '        'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    '        sprCodeTable.CtlRefresh()
    '        sprCodeTable.Col = mcSPCol_Code : sprCodeTable.Col2 = mcSPCol_Paid
    '        sprCodeTable.Row = rRow : sprCodeTable.Row2 = -1
    '        sprCodeTable.BlockMode = True
    '        sprCodeTable.Lock = False
    '        '        :      sprSpread.BackColor = Me.BackColor
    '        sprCodeTable.BlockMode = False

    '        sprCodeTable.SelectBlockOptions = SS_SELBLOCKOPT_ROWS
    '        sprCodeTable.UserResize = FPSpreadADO.UserResizeConstants.UserResizeColumns
    '        ''                sprCodeTable.MaxRows = sprCodeTable.DataRowCnt

    '    End If


    'End Sub
    Private Function InsretCheckedCell(ByVal rRow As Integer, Optional ByRef blnDelRec As Boolean = False) As Short
        Dim sSql As String = String.Empty
        Dim where As String = String.Empty

        Dim fn() As Object
        Dim XSource As New VBtoRecordSet()
        Dim strEx As String = String.Empty

        Dim strlen As Short
        strEx = String.Empty

        InsretCheckedCell = GetTypeLen()
        If mcs_Type = GetType_Renamed() And mcs_UIType = lblHidden.Text Then
            ' sprCodeTable.Col = 3
            sSql = "Select EXCLUDE from " & sTblSrc
            ' XSource.ActiveConnection.ConnectionString = Helper.GetConnectionString()
            XSource.OpenRecordset(sSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

            If sprCodeTable.Columns(0).Index = "1" And blnDelRec = False Then



                'sprCodeTable.Row = rRow
                'sprCodeTable.Col = 2
                If sprCodeTable.Text <> "" Then

                    If XSource.RecordCount > 0 Then
                        strEx = "X,T" & XstrTbl
                        strEx = strEx & "," & sprCodeTable.Text

                        If InStr(Len(strEx), strEx, ",") > 0 Then

                            strEx = VB.Left(strEx, Len(strEx) - 1)
                        End If
                        If InStr(strEx, ",,") > 0 Then
                            strEx = Replace(strEx, ",,", ",")
                        End If

                    End If
                    ReDim fn(0)
                    'UPGRADE_WARNING: Couldn't resolve default property of object fn(GetOrdinal()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    fn(GetOrdinal(XSource, "EXCLUDE")) = strEx
                    DBBeginTran()
                    DBUpdateRec(sTblSrc, XSource, fn, where)
                    DBCommit()

                    InsretCheckedCell = 2
                End If

            Else
                'sprCodeTable.Row = rRow
                'sprCodeTable.Col = 2
                If Trim(sprCodeTable.Text) <> "" Then

                    If XSource.RecordCount > 0 Then
                        strEx = "X,T" & XstrTbl
                        If InStr(strEx, Trim(sprCodeTable.Text)) > 0 Then
                            strlen = InStr(strEx, Trim(sprCodeTable.Text))
                            strEx = Replace(strEx, Trim(sprCodeTable.Text), Space(0))
                        End If
                        If InStr(Len(strEx), strEx, ",") > 0 Then

                            strEx = VB.Left(strEx, Len(strEx) - 1)
                        End If
                        If InStr(strEx, ",,") > 0 Then
                            strEx = Replace(strEx, ",,", ",")
                        End If

                    End If
                    ReDim fn(0)
                    'UPGRADE_WARNING: Couldn't resolve default property of object fn(GetOrdinal()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    fn(GetOrdinal(XSource, "EXCLUDE")) = strEx
                    '  DBBeginTran()
                    DBUpdateRec(sTblSrc, XSource, fn, where)
                    ' DBCommit()

                    InsretCheckedCell = 1
                End If

            End If

        End If


    End Function
    Private Function chkExclude(ByVal nRow As Integer, ByVal nCol As Integer) As Boolean

        Dim EXstrArr() As String
        Dim strsearch As String = String.Empty

        Dim strSQL As String = String.Empty

        Dim ctrl As Short
        Dim Xstatus As New VBtoRecordSet()



        'sprCodeTable.Row = nRow
        ' sprCodeTable.Col = nCol
        If Trim(sprCodeTable.Text) <> "" Then



            Xstr = Fetch_Exclude()
            strsearch = Xstr
            strsearch = Replace(strsearch, "'", "")

            If InStr(strsearch, sprCodeTable.Text) > 0 Then
                strSQL = String.Empty
                strSQL = "Select X_PSTATUS from property where "
                strSQL = strSQL & "X_PSTATUS = " & "'" & Trim(sprCodeTable.Text) & "'"

                Xstatus = New VBtoRecordSet()
                ' Xstatus.ActiveConnection.ConnectionString = Helper.GetConnectionString()
                Xstatus.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

                If Not Xstatus.EOF Then
                    '  LockSpread(True, nRow)
                    '  MsgBox strsearch & " cannot be deleted as this is used in system for Payoff transactions"

                    Exit Function
                End If
            End If

        End If




    End Function


    Private Sub sprCodeTable_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles sprCodeTable.CellValueChanged
        'Dim sWhere As String
        'If e.ColumnIndex = 0 Then

        '    For row As Integer = 0 To sprCodeTable.Rows.Count - 2

        '        If sprCodeTable.Rows(row).Cells(0).Value IsNot Nothing AndAlso row <> e.RowIndex AndAlso sprCodeTable.Rows(row).Cells(0).Value.Equals(sprCodeTable.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then

        '            Dim cellvalue As String = sprCodeTable.CurrentRow.Cells(0).Value
        '            MessageBox.Show("Code" + " " + cellvalue + " " + "already exists. Please try another!")
        '            'LoadActSpread()
        '        Else

        '        End If
        '    Next
        'End If
        'If IsDBNull(sprCodeTable.CurrentCell.FormattedValue) Then
        '    strnew = String.Empty
        'Else
        '    strnew = sprCodeTable.CurrentCell.FormattedValue.ToString()
        'End If

        'If String.IsNullOrEmpty(strnew) AndAlso e.ColumnIndex = 0 Then
        '    sprCodeTable.CurrentCell.Value = strold

        'Else
        '    If StrComp(strold, strnew, CompareMethod.Text) <> 0 Then

        '        sSql = "Select ACTIONCODE, ACTDESC from ACTIONCODES ORDER BY ACTIONCODE"
        '        ' dsSource.ActiveConnection.ConnectionString = Helper.GetConnectionString()
        '        dsSource.OpenRecordset(sSql, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        '        Dim aGLAcct(dsSource.ColumnCount - 1) As Object

        '        Dim Row As Integer = sprCodeTable.CurrentRow.Cells(0).Value
        '        aGLAcct(GetOrdinal(dsSource, "ACTIONCODE")) = sprCodeTable.CurrentRow.Cells(0).Value
        '        If IsDBNull(sprCodeTable.CurrentRow.Cells(1).Value) Then
        '            aGLAcct(GetOrdinal(dsSource, "ACTDESC")) = String.Empty
        '        Else
        '            aGLAcct(GetOrdinal(dsSource, "ACTDESC")) = sprCodeTable.CurrentRow.Cells(1).Value
        '        End If
        '        'sWhere = "ACTIONCODE = '" & Row & "'" & "' AND "
        '        sWhere = "ACTIONCODE ='" & actionCode & "'"
        '        DBUpdateRec("ACTIONCODES", dsSource, aGLAcct, sWhere)
        '        ' LoadActSpread()
        '        sprCodeTable.RefreshEdit()
        '    End If
        'End If



    End Sub
    Private Sub sprCodeTable_EditingControlShowing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles sprCodeTable.EditingControlShowing
        Dim a As String = Me.sprCodeTable.CurrentCell.Value
        If (Me.sprCodeTable.CurrentCell.ColumnIndex = 0) Then
            AddHandler e.Control.KeyPress, AddressOf cell_KeyDown

        End If
        If (Me.sprCodeTable.CurrentCell.ColumnIndex = 1) Then
            AddHandler e.Control.KeyPress, AddressOf cell_KeyDown

        End If
    End Sub
    Private Sub cell_KeyDown(ByVal sender As Object, ByVal e As KeyPressEventArgs)

        If (Me.sprCodeTable.CurrentCell.ColumnIndex = 1) Then

            Dim currentValue As String = Me.sprCodeTable.CurrentCell.GetEditedFormattedValue(Me.sprCodeTable.CurrentRow.Index, DataGridViewDataErrorContexts.Display).ToString()
            If currentValue.Length = 2 AndAlso e.KeyChar <> Convert.ToChar(Keys.Back) AndAlso e.KeyChar <> Convert.ToChar(Keys.Delete) Then
                e.Handled = True
            Else
                If Not (Char.IsDigit(e.KeyChar)) AndAlso e.KeyChar <> ChrW(Keys.Back) AndAlso e.KeyChar <> ChrW(Keys.Delete) Then
                    e.Handled = True
                End If
            End If

        End If
        If (Me.sprCodeTable.CurrentCell.ColumnIndex = 0) Then
            Dim currentValue As String = Me.sprCodeTable.CurrentCell.GetEditedFormattedValue(Me.sprCodeTable.CurrentRow.Index, DataGridViewDataErrorContexts.Display).ToString()
            If currentValue.Length = 50 AndAlso e.KeyChar <> Convert.ToChar(Keys.Back) AndAlso e.KeyChar <> Convert.ToChar(Keys.Delete) Then
                e.Handled = True
            End If
        End If

    End Sub
    Private Sub sprCodeTable_CellBeginEdit(ByVal sender As Object, ByVal e As DataGridViewCellCancelEventArgs) Handles sprCodeTable.CellBeginEdit
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Exit Sub
        If IsDBNull(sprCodeTable.CurrentCell.Value) Then
            strold = String.Empty
        Else
            strold = sprCodeTable.CurrentCell.Value
        End If

        actionCode = sprCodeTable.Rows(e.RowIndex).Cells(0).Value
        'strold = sprCodeTable.CurrentCell.FormattedValue.ToString()
        ' Me.sprCodeTable.Tag = Me.sprCodeTable.CurrentCell.Value
    End Sub

    Private Sub sprCodeTable_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles sprCodeTable.CellMouseDown
        MessageBox.Show("fired")
    End Sub
End Class

