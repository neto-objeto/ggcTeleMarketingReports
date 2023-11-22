Imports System.Windows.Forms
Imports ggcAppDriver
Imports ADODB
Imports MySql.Data.MySqlClient
Public Class frmMCSalesCriteria

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


    Private Sub cmdButton01_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdButton01.Click

        

        If Not (IsDate(txtField01.Text) And
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

    Private Sub frmMCSalesCriteria_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        txtField01.Text = Format(Now(), "yyyy-MM-dd")
        txtField02.Text = txtField01.Text
        txtField03.Text = txtField03.Text
        txtField04.Text = txtField04.Text
    End Sub


    Private Sub txtField_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtField03.KeyDown, txtField04.KeyDown
        If e.KeyCode = Keys.F3 Or e.KeyCode = Keys.Return Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)

            Dim loIndex As Integer
            loIndex = Val(Mid(loTxt.Name, 9))

            If Mid(loTxt.Name, 1, 8) = "txtField" Then
                Select Case loIndex
                    Case 3
                        If (txtField03.Text <> "") Then SearchAgent(loTxt.Text, False, True)

                    Case 4
                        If (txtField04.Text <> "") Then SearchBranch(loTxt.Text, False, True)

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

    Public Function SearchBranch(ByVal fsValue As String _
                            , ByVal fbByCode As Boolean _
                            , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sBranchCD LIKE " & strParm("%" & fsValue)
        Else
            lsFilter = "c.sBranchNm like " & strParm(fsValue & "%")
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" &
                       "  a.sBranchCD" &
                       ", a.sBranchNm" &
               " FROM Branch a" &
                   " LEFT JOIN Branch_Others b " &
                   " ON a.sBranchCD = b.sBranchCD " &
               " WHERE b.cDivision = '1' "
        IIf(fbByCode = False, " AND a.cRecdStat = '1'", "")

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oDriver _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sBranchCD»sBranchNm" _
                                             , "Code»Branch", _
                                             , "a.sBranchCD»a.sBranchNm" _
                                             , IIf(fbByCode, 0, 1))
            If IsNothing(loRow) Then
                Return False
            Else
                txtField04.Text = (loRow.Item("sBranchNm"))
                txtField04.Tag = (loRow.Item("sBranchCD"))
         
            End If
        End If
        Return True
    End Function


End Class
