Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports CrystalDecisions.CrystalReports.Engine

Public Class clsCallStatusReport
    Private p_oDriver As ggcAppDriver.GRider
    Private p_oSTRept As DataSet
    Private p_oDTSrce As DataTable

    Private p_nReptType As Integer      '0=Summary;1=Detail
    Private p_abInclude(4) As Integer   '0=Default;1=Financing
    Private p_dDateFrom As Date
    Private p_dDateThru As Date
    Private p_sAgentID As String
    Private p_sBranchCD As String


    Public Function getParameter() As Boolean

        Dim loFrm As frmCallStatusCriteria
        loFrm = New frmCallStatusCriteria
        loFrm.GRider = p_oDriver

        loFrm.ShowDialog()

        If loFrm.isOkey Then
            'Since we have not allowed the report type to be edited
            p_nReptType = 0

            p_abInclude(0) = loFrm.chkInclude00.Checked
            p_abInclude(1) = loFrm.chkInclude01.Checked
            p_abInclude(2) = loFrm.chkInclude02.Checked
            p_abInclude(3) = loFrm.chkInclude03.Checked
            p_abInclude(4) = loFrm.chkInclude04.Checked

            p_dDateFrom = loFrm.txtField01.Text
            p_dDateThru = loFrm.txtField02.Text
            p_sAgentID = loFrm.txtField03.Tag
            Debug.Print(p_sAgentID)
            loFrm = Nothing
            Return True
        Else
            loFrm = Nothing
            Return False
        End If
    End Function

    Public Function ReportTrans() As Boolean
        Dim oProg As frmProgress

        Dim lsSQL As String 'whole statement
        Dim lsQuery1 As String


        'Show progress bar
        oProg = New frmProgress
        oProg.PistonInfo = p_oDriver.AppPath & "/piston.avi"
        oProg.ShowTitle("EXTRACTING RECORDS FROM DATABASE")
        oProg.ShowProcess("Please wait...")
        oProg.Show()

        lsQuery1 = "SELECT CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNULL(b.sSuffixNm, '') = '', '', CONCAT(' ', b.sSuffixNm)), IF(IFNULL(b.sMiddName, '') = '', '', CONCAT(' ', b.sMiddName))) `sClientNm`" &
                        ", a.sMobileNo `sMobileNo`" &
                        ", a.sRemarksx `sRemarksx`" &
                        ", a.sApprovCd `sApprovCd`" &
                        ", CASE a.cTranStat" &
                            " WHEN '0' THEN 'OPEN'" &
                            " WHEN '1' THEN 'QUEUED'" &
                            " WHEN '2' THEN 'CALLED'" &
                            " WHEN '3' THEN 'DISCARDED'" &
                            " ELSE 'RECYCLED'" &
                            " END `cTranStat`" &
                        ", a.sSourceCD `sSourceCD`" &
                        ", a.sModified `sModified`" &
                        ", a.dModified `dModified`" &
                        ", a.sAgentIDx `sAgentIDx`" &
                " FROM Call_Outgoing a" &
                    " LEFT JOIN Client_Master b" &
                        " ON a.sClientID = b.sClientID" &
                " WHERE a.dModified BETWEEN " & strParm(Format(p_dDateFrom, "yyyy-MM-dd") & " 00:00:00") &
                        " AND " & strParm(Format(p_dDateThru, "yyyy-MM-dd") & " 23:59:00")



        If p_sAgentID <> "" Then
            lsQuery1 = AddCondition(lsQuery1, "a.sAgentIDx = " & strParm(p_sAgentID))
        End If

        lsSQL = ""

        lsSQL = lsQuery1

        If p_abInclude(1) And lsSQL <> "" Then  'open 
            lsSQL = AddCondition(lsSQL, "a.cTranStat = '0'")
        ElseIf p_abInclude(2) And lsSQL <> "" Then  'queued 
            lsSQL = AddCondition(lsSQL, "a.cTranStat = '1'")
        ElseIf p_abInclude(3) And lsSQL <> "" Then  'called  & recycled
            lsSQL = AddCondition(lsSQL, "a.cTranStat IN ('2','5')")
        ElseIf p_abInclude(4) And lsSQL <> "" Then  'discarded
            lsSQL = AddCondition(lsSQL, "a.cTranStat = '3'")
        End If

        If lsSQL <> "" Then
            lsSQL = lsSQL & " ORDER BY sAgentIDx, dModified ASC"
        End If
        Debug.Print(lsSQL)

        p_oDTSrce = p_oDriver.ExecuteQuery(lsSQL)

        Dim loDtaTbl As DataTable = getRptTable()
        Dim lnCtr As Integer

        oProg.ShowTitle("LOADING RECORDS")
        oProg.MaxValue = p_oDTSrce.Rows.Count

        For lnCtr = 0 To p_oDTSrce.Rows.Count - 1

            oProg.ShowProcess("Loading " & p_oDTSrce(lnCtr).Item("sClientNm") & "...")

            loDtaTbl.Rows.Add(addRow(lnCtr, loDtaTbl))
        Next

        oProg.ShowSuccess()

        Dim clsRpt As clsReports
        clsRpt = New clsReports
        clsRpt.GRider = p_oDriver
        'Set the Report Source Here
        If Not clsRpt.initReport("TLMC2") Then
            Return False
        End If

        Dim loRpt As ReportDocument = clsRpt.ReportSource

        Dim loTxtObj As CrystalDecisions.CrystalReports.Engine.TextObject
        loTxtObj = loRpt.ReportDefinition.Sections(0).ReportObjects("txtCompany")
        loTxtObj.Text = p_oDriver.BranchName

        'Set Branch Address
        loTxtObj = loRpt.ReportDefinition.Sections(0).ReportObjects("txtAddress")
        loTxtObj.Text = p_oDriver.Address & vbCrLf & p_oDriver.TownCity & " " & p_oDriver.ZippCode & vbCrLf & p_oDriver.Province

        'Set First Header
        loTxtObj = loRpt.ReportDefinition.Sections(1).ReportObjects("txtHeading1")
        loTxtObj.Text = "Lead's Report"

        'Set Second Header
        loTxtObj = loRpt.ReportDefinition.Sections(1).ReportObjects("txtHeading2")
        loTxtObj.Text = Format(p_dDateFrom, "MMMM dd yyyy") & " to " & Format(p_dDateThru, "MMMM dd yyyy")

        loTxtObj = loRpt.ReportDefinition.Sections(3).ReportObjects("txtRptUser")
        loTxtObj.Text = Decrypt(p_oDriver.UserName, "08220326")

        loRpt.SetDataSource(p_oSTRept)
        clsRpt.showReport()

        Return True
    End Function

    Private Function getRptTable() As DataTable
        'Initialize DataSet
        p_oSTRept = New DataSet

        'Load the data structure of the Dataset
        'Data structure was saved at DataSet1.xsd 
        p_oSTRept.ReadXmlSchema(p_oDriver.AppPath & "\vb.net\Reports\DataSet1.xsd")

        'Return the schema of the datatable derive from the DataSet 
        Return p_oSTRept.Tables(0)
    End Function

    Private Function addRow(ByVal lnRow As Integer, ByVal foSchemaTable As DataTable) As DataRow
        'ByVal foDTInclue As DataTable
        Dim loDtaRow As DataRow

        'Create row based on the schema of foSchemaTable
        loDtaRow = foSchemaTable.NewRow

        loDtaRow.Item("nField01") = lnRow + 1
        loDtaRow.Item("sField01") = p_oDTSrce(lnRow).Item("sClientNm")
        loDtaRow.Item("sField02") = p_oDTSrce(lnRow).Item("sMobileNo")
        loDtaRow.Item("sField03") = p_oDTSrce(lnRow).Item("cTranStat")
        loDtaRow.Item("sField04") = getAgent(p_oDTSrce(lnRow).Item("sAgentIDx"))
        Return loDtaRow
    End Function

    Private Function getAgent(ByVal sAgentIDx As String) As String
        Dim lsSQL As String

        lsSQL = "SELECT sUserName FROM xxxSysUser WHERE sUserIDxx = " & strParm(sAgentIDx)

        Dim loDT As DataTable
        loDT = p_oDriver.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then Return ""

        Return Decrypt(loDT(0)("sUserName"), "08220326")
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oDriver = foRider
        p_oSTRept = Nothing
        p_oDTSrce = Nothing
    End Sub


End Class