Imports System.Windows.Forms
Imports ggcAppDriver
Imports ADODB
Imports MySql.Data.MySqlClient
Public Class frmCallStatusCriteria

    Private pb_ChkdOk As Boolean
    Private pn_Loaded As Integer
    Private p_oDriver As ggcAppDriver.GRider

    Public WriteOnly Property GRider() As ggcAppDriver.GRider
        Set(ByVal foValue As ggcAppDriver.GRider)
            p_oDriver = foValue
        End Set
    End Property

    Public Function isOkey() As Boolean
        Return pb_ChkdOk
    End Function

    Private Sub chkInclude_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkInclude00.CheckedChanged, chkInclude01.CheckedChanged _
            , chkInclude02.CheckedChanged, chkInclude03.CheckedChanged, chkInclude04.CheckedChanged

        Dim lochkBox As CheckBox
        lochkBox = CType(sender, System.Windows.Forms.CheckBox)

        If Mid(lochkBox.Name, 1, 10) = "chkInclude" Then
            Dim loIndex As Integer
            loIndex = Val(Mid(lochkBox.Name, 12))

            ' Disconnect event handlers temporarily
            For i As Integer = 0 To 4
                Dim chkBox As CheckBox = TryCast(gbxPanel03.Controls("chkInclude0" & i), CheckBox)

                ' Check if the control exists before accessing its properties
                If chkBox IsNot Nothing Then
                    RemoveHandler chkBox.CheckedChanged, AddressOf chkInclude_CheckedChanged
                    chkBox.Checked = (i = loIndex)
                    AddHandler chkBox.CheckedChanged, AddressOf chkInclude_CheckedChanged
                End If
            Next i
        End If
    End Sub


    Private Sub cmdButton01_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdButton01.Click

        If Not (chkInclude00.Checked Or
                chkInclude01.Checked Or
                chkInclude02.Checked Or
                chkInclude03.Checked Or
                chkInclude04.Checked) Then
            MsgBox("There are items selected in the INCLUDE group." & vbCrLf &
                   "Please check your entry try again!", vbOKOnly, "Parameter Validation")
            Exit Sub

        ElseIf Not (IsDate(txtField01.Text) And
                    IsDate(txtField02.Text)) Then

            MsgBox("There are invalid date in the RANGE group." & vbCrLf &
                   "Please check your entry try again!", vbOKOnly, "Parameter Validation")
            Exit Sub

        ElseIf CDate(txtField01.Text) > CDate(txtField02.Text) Then
            MsgBox("FROM parameter seems to be higher than THRU in the RANGE group." & vbCrLf &
                   "Please check your entry try again!", vbOKOnly, "Parameter Validation")
            Exit Sub
        End If

        pb_ChkdOk = True

        Me.Hide()
    End Sub

    Private Sub cmdButton00_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdButton00.Click
        pb_ChkdOk = False
        Me.Hide()
    End Sub

    Private Sub txtField01_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtField01.Validated
        If IsDate(txtField01.Text) Then
            txtField01.Text = Format(CDate(txtField01.Text), "yyyy-MM-dd")
        Else
            txtField01.Text = Format(Now(), "yyyy-MM-dd")
        End If
    End Sub

    Private Sub txtField02_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtField02.Validated
        If IsDate(txtField02.Text) Then
            txtField02.Text = Format(CDate(txtField02.Text), "yyyy-MM-dd")
        Else
            txtField02.Text = Format(Now(), "yyyy-MM-dd")
        End If
    End Sub

    Private Sub frmCallStatusCriteria_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        txtField01.Text = Format(Now(), "yyyy-MM-dd")
        txtField02.Text = txtField01.Text
        txtField03.Text = txtField03.Text
        chkInclude00.Checked = True
        chkInclude01.Checked = False
        chkInclude02.Checked = False
        chkInclude03.Checked = False
        chkInclude04.Checked = False
    End Sub


    Private Sub txtField_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtField03.KeyDown
        If e.KeyCode = Keys.F3 Or e.KeyCode = Keys.Return Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)

            Dim loIndex As Integer
            loIndex = Val(Mid(loTxt.Name, 9))

            If Mid(loTxt.Name, 1, 8) = "txtField" Then
                Select Case loIndex
                    Case 3
                        If (txtField03.Text <> "") Then SearchAgent(loTxt.Text, False, True)

                End Select
            End If
        End If
    End Sub

    Public Function SearchAgent(
                    ByVal fsValue As String _
                  , ByVal fbByCode As Boolean _
                    , ByVal fbIsSrch As Boolean) As Boolean

        Dim lsSQL As String

        'Initialize SQL filter
        lsSQL = getSQ_Agent()


        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sUserIDxx LIKE " & strParm("%" & fsValue)
        Else
            lsFilter = "c.sCompnyNm like " & strParm(fsValue & "%")
        End If
        If fbIsSrch Then
            Debug.Print(lsSQL)
            Dim loDta As DataRow = KwikSearch(p_oDriver _
                                            , lsSQL _
                                            , fbIsSrch _
                                             , fsValue _
                                             , "sCompnyNm»sUserIDxx" _
                                             , "Company Name»Agent ID", _
                                             , "a.sUserIDxx»c.sCompnyNm")
            If IsNothing(loDta) Then

                Return False
            Else
                txtField03.Text = (loDta.Item("sCompnyNm"))
                txtField03.Tag = (loDta.Item("sUserIDxx"))

            End If
        End If
        Return True
    End Function

    Private Function getSQ_Agent() As String
        Return "SELECT a.sUserIDxx sUserIDxx" &
                    ", c.sCompnyNm sCompnyNm" &
              " FROM xxxSysUser  a " &
              " LEFT JOIN Employee_Master001 b " &
              " ON a.sEmployNo = b.sEmployID " &
              " LEFT JOIN Client_Master c " &
              " ON b.sEmployID = c.sClientID " &
              " WHERE sProdctID = 'TeleMktg' "

    End Function
End Class
