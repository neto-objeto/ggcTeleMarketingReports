Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports CrystalDecisions.CrystalReports.Engine

Public Class clsCallReport
    Private p_oDriver As ggcAppDriver.GRider
    Private p_oSTRept As DataSet
    Private p_oDTSrce As DataTable

    Private p_nReptType As Integer      '0=Summary;1=Detail
    Private p_abInclude(1) As Integer   '0=Default;1=Financing
    Private p_dDateFrom As Date
    Private p_dDateThru As Date
    Private p_sBranchCD As String

    Public Function getParameter() As Boolean
        Dim loFrm As frmOutboundCall

        loFrm = New frmOutboundCall

        'Disable Report Type Group
        loFrm.gbxPanel01.Enabled = False
        'Set Detail as Report Type
        loFrm.rbtTypex02.Checked = True

        loFrm.ShowDialog()

        If loFrm.isOkey Then
            'Since we have not allowed the report type to be edited
            p_nReptType = 0

            p_abInclude(0) = loFrm.chkInclude01.Checked
            p_abInclude(1) = loFrm.chkInclude02.Checked

            p_dDateFrom = loFrm.txtField01.Text
            p_dDateThru = loFrm.txtField02.Text

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
        Dim lsQuery2 As String

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
                            " WHEN '2' THEN a.cTLMStatx" &
                            " ELSE CONCAT(a.cTLMStatx, '(RECYCLED)')" &
                            " END `cTLMStatx`" &
                        ", a.sSourceCD `sSourceCD`" &
                        ", a.sModified `sModified`" &
                        ", a.dModified `dModified`" &
                        ", a.sAgentIDx `sAgentIDx`" &
                " FROM Call_Outgoing a" &
                    " LEFT JOIN Client_Master b" &
                        " ON a.sClientID = b.sClientID" &
                " WHERE a.sSourceCd <> 'LEND'" &
                    " AND a.cTranStat IN ('2')" &
                    " AND a.dModified BETWEEN " & strParm(Format(p_dDateFrom, "yyyy-MM-dd") & " 00:00:00") &
                        " AND " & strParm(Format(p_dDateThru, "yyyy-MM-dd") & " 23:59:00")

        lsQuery2 = "SELECT CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNULL(b.sSuffixNm, '') = '', '', CONCAT(' ', b.sSuffixNm)), IF(IFNULL(b.sMiddName, '') = '', '', CONCAT(' ', b.sMiddName))) `sClientNm`" &
                        ", a.sMobileNo `sMobileNo`" &
                        ", a.sRemarksx `sRemarksx`" &
                        ", a.sApprovCd `sApprovCd`" &
                        ", CASE a.cTranStat" &
                            " WHEN '2' THEN a.cTLMStatx" &
                            " ELSE CONCAT(a.cTLMStatx, '(RECYCLED)')" &
                            " END `cTLMStatx`" &
                        ", a.sSourceCD `sSourceCD`" &
                        ", a.sModified `sModified`" &
                        ", a.dModified `dModified`" &
                        ", a.sAgentIDx `sAgentIDx`" &
               " FROM Call_Outgoing a" &
                   " LEFT JOIN Client_Master b" &
                       " ON a.sClientID = b.sClientID" &
               " WHERE a.sSourceCd = 'LEND'" &
                   " AND a.cTranStat IN ('2')" &
                   " AND a.dModified BETWEEN " & strParm(Format(p_dDateFrom, "yyyy-MM-dd") & " 00:00:00") &
                        " AND " & strParm(Format(p_dDateThru, "yyyy-MM-dd") & " 23:59:00")


        Dim lsAgent As String = p_oDriver.UserID
        If lsAgent = "M0T1160004" Or
            lsAgent = "M001160022" Or
            lsAgent = "M001160024" Or
            lsAgent = "M001111122" Or
            lsAgent = "M001180036" Or
            lsAgent = "M0T1230002" Then
            lsAgent = ""
        End If

        If lsAgent <> "" Then
            lsQuery1 = AddCondition(lsQuery1, "a.sAgentIDx = " & strParm(lsAgent))
            lsQuery2 = AddCondition(lsQuery2, "a.sAgentIDx = " & strParm(lsAgent))
        End If

        lsSQL = ""
        If p_abInclude(0) Then
            lsSQL = lsQuery1
        End If

        If p_abInclude(1) And lsSQL <> "" Then
            lsSQL = lsSQL & " UNION " & lsQuery2
        ElseIf p_abInclude(1) Then
            lsSQL = lsQuery2
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
        If Not clsRpt.initReport("TLMC1") Then
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
        loTxtObj.Text = "Outbound Call Report"

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
        loDtaRow.Item("sField03") = p_oDTSrce(lnRow).Item("sApprovCd")
        loDtaRow.Item("sField04") = p_oDTSrce(lnRow).Item("sSourceCd")
        loDtaRow.Item("sField05") = IIf(p_oDTSrce(lnRow).Item("cTLMStatx") = "UR", "CR", p_oDTSrce(lnRow).Item("cTLMStatx"))
        loDtaRow.Item("sField06") = Left(p_oDTSrce(lnRow).Item("sRemarksx"), 72)
        loDtaRow.Item("sField07") = getAgent(p_oDTSrce(lnRow).Item("sAgentIDx"))
        loDtaRow.Item("sField08") = Format(p_oDTSrce(lnRow).Item("dModified"), "MM-dd-yyyy  hh:mm")
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