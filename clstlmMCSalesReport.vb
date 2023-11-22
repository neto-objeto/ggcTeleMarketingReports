Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports CrystalDecisions.CrystalReports.Engine

Public Class clstlmMCSalesReport
    Private p_oDriver As ggcAppDriver.GRider
    Private p_oSTRept As DataSet
    Private p_oDTSrce As DataTable

    Private p_nReptType As Integer
    Private p_dDateFrom As Date
    Private p_dDateThru As Date
    Private p_sAgentID As String
    Private p_sBranchCD As String


    Public Function getParameter() As Boolean

        Dim loFrm As frmMCSalesCriteria
        loFrm = New frmMCSalesCriteria
        loFrm.GRider = p_oDriver

        loFrm.ShowDialog()

        If loFrm.isOkey Then
            'Since we have not allowed the report type to be edited
            p_nReptType = 0

            p_dDateFrom = loFrm.txtField01.Text
            p_dDateThru = loFrm.txtField02.Text
            p_sAgentID = loFrm.txtField03.Tag
            p_sBranchCD = loFrm.txtField04.Tag
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

        lsQuery1 = " SELECT  " &
                    "   CONCAT(d.sLastName, ', ', d.sFrstName, IF(IFNULL(d.sSuffixNm, '') = '', '', CONCAT(' ', d.sSuffixNm)), IF(IFNULL(d.sMiddName, '') = '', '', CONCAT(' ', d.sMiddName))) sClientNm " &
                    " , a.sMobileNo sMobileNo " &
                    " , b.sAgentIDx sAgentIDx " &
                    " , c.sDRNoxxxx sDRNoxxxx " &
                    " , h.sBrandNme sBrandNme " &
                    " , g.sModelNme sModelNme " &
                    " , i.sBranchNm sBranchNm " &
                    " FROM TLM_Sales a " &
                    " LEFT JOIN Call_Outgoing b " &
                    " ON a.sSourceNo = b.sTransNox " &
                    " LEFT JOIN MC_SO_Master c " &
                    " ON a.sReferNox = c.sTransNox " &
                    " LEFT JOIN Client_Master d " &
                    " ON c.sClientID = d.sClientID " &
                    " LEFT JOIN MC_SO_Detail e " &
                    " ON c.sTransNox = e.sTransNox  " &
                    " LEFT JOIN MC_Serial f " &
                    " ON e.sSerialID = f.sSerialID " &
                    " LEFT JOIN MC_Model g " &
                    " ON  f.sModelIDx =  g.sModelIDx " &
                    " LEFT JOIN Brand h " &
                    " ON g.sBrandIDx = h.sBrandIDx " &
                    " LEFT JOIN Branch i " &
                    " ON  f.sBranchCd = i.sBranchCd " &
                " WHERE a.dModified BETWEEN " & strParm(Format(p_dDateFrom, "yyyy-MM-dd") & " 00:00:00") &
                        " AND " & strParm(Format(p_dDateThru, "yyyy-MM-dd") & " 23:59:00")



        If p_sAgentID <> "" Then
            lsQuery1 = AddCondition(lsQuery1, "b.sAgentIDx = " & strParm(p_sAgentID))
        End If
        If p_sBranchCD <> "" Then
            lsQuery1 = AddCondition(lsQuery1, "i.sBranchCd = " & strParm(p_sBranchCD))
        End If

        lsSQL = ""

        lsSQL = lsQuery1

       

        If lsSQL <> "" Then
            lsSQL = lsSQL & " ORDER BY sAgentIDx, a.dModified ASC"
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
        If Not clsRpt.initReport("TLMC3") Then
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
        loTxtObj.Text = "MC Sales Report"

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
        loDtaRow.Item("sField03") = getAgent(p_oDTSrce(lnRow).Item("sAgentIDx"))
        loDtaRow.Item("sField04") = p_oDTSrce(lnRow).Item("sDRNoxxxx")
        loDtaRow.Item("sField05") = p_oDTSrce(lnRow).Item("sBrandNme")
        loDtaRow.Item("sField06") = p_oDTSrce(lnRow).Item("sModelNme")
        loDtaRow.Item("sField07") = p_oDTSrce(lnRow).Item("sBranchNm")
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