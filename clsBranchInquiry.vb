Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports CrystalDecisions.CrystalReports.Engine

Public Class clsBranchInquiry
    Private p_oDriver As ggcAppDriver.GRider
    Private p_oSTRept As DataSet
    Private p_oDTSrce As DataTable

    Private p_nReptType As Integer
    Private p_dDateFrom As Date
    Private p_dDateThru As Date
    Private p_sAreaCd As String
    Private p_sBranchCD As String


    Public Function getParameter() As Boolean

        Dim loFrm As frmBranchCriteria
        loFrm = New frmBranchCriteria
        loFrm.GRider = p_oDriver

        loFrm.ShowDialog()

        If loFrm.isOkey Then
            'Since we have not allowed the report type to be edited
            p_nReptType = 0

            p_dDateFrom = loFrm.txtField01.Text
            p_dDateThru = loFrm.txtField02.Text
            p_sAreaCd = loFrm.txtField03.Tag
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

        lsQuery1 = " SELECT " &
                    " h.sAreaDesc Area " &
                    " , e.sBranchNm Branch " &
                    " , CASE a.sInquiryx " &
                        " WHEN 'WI' THEN 'Walk In' " &
                        " WHEN 'FB' THEN 'Facebook' " &
                        " WHEN 'WS' THEN 'Website' " &
                        " WHEN 'BF' THEN 'Biyaheng Fiesta' " &
                        " WHEN 'DS' THEN 'MC Display' " &
                        " WHEN 'FS' THEN 'Free Service Check Up' " &
                        " WHEN 'BR' THEN 'Branch Request' " &
                        " WHEN 'ER' THEN 'Employee Referral' " &
                        " WHEN 'MD' THEN 'Management Discount' " &
                        " WHEN 'SR' THEN 'Suzuki Referral' " &
                        " END `Inquiry Type` " &
                    " , a.dTransact `Date of Inquiry` " &
                    " , b.sCompnyNm `Customer Name` " &
                    " , b.sMobileNo `Mobile No` " &
                    " , CASE i.cSubscrbr " &
                        " WHEN '0' THEN 'GLOBE' " &
                        " WHEN '1' THEN 'SMART' " &
                        " WHEN '2' THEN 'SUN' " &
                        " WHEN '3' THEN 'DITO' " &
                        " ELSE'NOT SET' " &
                        " END Network " &
                    " , c.sModelNme Model " &
                    " , d.sColorNme Color " &
                    " , a.dFollowUp `Follow Up` " &
                    " , a.dTargetxx Target " &
                    " , a.dPurchase Purchased " &
                    " , CASE a.cTranStat " &
                        " WHEN '0' THEN 'OPEN' " &
                        " WHEN '1' THEN 'SCHEDULED' " &
                        " WHEN '2' THEN 'PURCHASED' " &
                        " WHEN '3' THEN 'CANCELLED' " &
                        " WHEN '4' THEN 'HAS DUPLICATE LEADS' " &
                        " WHEN '5' THEN 'RECYCLED' " &
                        " ELSE 'UNKNOWN' " &
                        " END `Inquiry Status` " &
                    " , f.dTransact `Date added to TLM leads` " &
                    " , f.cTLMStatx `TLM Stat` " &
                    " , f.sRemarksx Remarks " &
                    " , f.cSMSStatx `SMS Stat` " &
                    " , IF(f.dCallStrt = '1900-01-01', '(NULL)', f.dCallStrt) `Call Start` " &
                    " , IF(f.dCallEndx = '1900-01-01', '(NULL)', f.dCallEndx) `Call End` " &
                    " , f.dModified TimeStamp " &
                    " , a.sTransNox `Source No` " &
                    " FROM MC_Product_Inquiry a " &
                    " LEFT JOIN MC_Model c " &
                    " ON a.sModelIDx = c.sModelIDx " &
                    " LEFT JOIN Color d " &
                    " ON a.sColorIDx = d.sColorIDx " &
                    " LEFT JOIN Branch e " &
                    " ON LEFT(a.sTransNox, 4) = e.sBranchCd " &
                    " LEFT JOIN Branch_Others g " &
                    " ON e.sBranchCd = g.sBranchCd " &
                    " LEFT JOIN Branch_Area h " &
                    " ON g.sAreaCode = h.sAreaCode " &
                    " LEFT JOIN Call_Outgoing f " &
                    " ON a.sTransNox = f.sReferNox " &
                    " AND f.sSourceCd = 'Inqr' " &
                    " , Client_Master b " &
                    " LEFT JOIN Client_Mobile i " &
                    " ON b.sClientID = i.sClientID " &
                    " AND b.sMobileNo = i.sMobileNo " &
                    " WHERE a.sClientID = b.sClientID " &
                    " AND a.dTransact BETWEEN  " & strParm(Format(p_dDateFrom, "yyyy-MM-dd")) &
                                            " AND " & strParm(Format(p_dDateThru, "yyyy-MM-dd")) &
                    " GROUP BY a.sTransnox; " &
                    " ORDER BY f.dTransact DESC "

        If p_sAreaCd <> "" Then
            lsQuery1 = AddCondition(lsQuery1, "h.sAreaCode = " & strParm(p_sAreaCd))
        End If
        If p_sBranchCD <> "" Then
            lsQuery1 = AddCondition(lsQuery1, "e.sBranchCd = " & strParm(p_sBranchCD))
        End If

        lsSQL = ""

        lsSQL = lsQuery1



        If lsSQL <> "" Then
            lsSQL = lsSQL & " ORDER BY h.sAreaCode, f.dModified ASC"
        End If
        Debug.Print(lsSQL)

        p_oDTSrce = p_oDriver.ExecuteQuery(lsSQL)

        If p_oDTSrce.Rows.Count <= 0 Then
            MsgBox("No Record Found! ")
            Return False
        End If

        Dim loDtaTbl As DataTable = getRptTable()
        Dim lnCtr As Integer

        oProg.ShowTitle("LOADING RECORDS")
        oProg.MaxValue = p_oDTSrce.Rows.Count

        For lnCtr = 0 To p_oDTSrce.Rows.Count - 1

            oProg.ShowProcess("Loading " & p_oDTSrce(lnCtr).Item("Customer Name") & "...")

            loDtaTbl.Rows.Add(addRow(lnCtr, loDtaTbl))
        Next

        oProg.ShowSuccess()

        Dim clsRpt As clsReports
        clsRpt = New clsReports
        clsRpt.GRider = p_oDriver
        'Set the Report Source Here
        If Not clsRpt.initReport("TLMC4") Then
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
        loTxtObj.Text = "Branch Inquiry"

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

        loDtaRow.Item("sField01") = p_oDTSrce(lnRow).Item("Area")
        loDtaRow.Item("sField02") = p_oDTSrce(lnRow).Item("Branch")
        loDtaRow.Item("sField03") = p_oDTSrce(lnRow).Item("Inquiry Type")
        loDtaRow.Item("sField04") = p_oDTSrce(lnRow).Item("Date of Inquiry")
        loDtaRow.Item("sField05") = p_oDTSrce(lnRow).Item("Customer Name")
        loDtaRow.Item("sField06") = p_oDTSrce(lnRow).Item("Mobile No")
        loDtaRow.Item("sField07") = p_oDTSrce(lnRow).Item("Network")
        loDtaRow.Item("sField08") = p_oDTSrce(lnRow).Item("Model")
        loDtaRow.Item("sField09") = p_oDTSrce(lnRow).Item("Color")
        loDtaRow.Item("sField10") = p_oDTSrce(lnRow).Item("Follow Up")
        loDtaRow.Item("sField11") = p_oDTSrce(lnRow).Item("Target")
        loDtaRow.Item("sField12") = p_oDTSrce(lnRow).Item("Purchased")
        loDtaRow.Item("sField13") = p_oDTSrce(lnRow).Item("Inquiry Status")
        loDtaRow.Item("sField14") = p_oDTSrce(lnRow).Item("Date added to TLM leads")
        loDtaRow.Item("sField15") = p_oDTSrce(lnRow).Item("TLM Stat")
        loDtaRow.Item("sField16") = p_oDTSrce(lnRow).Item("Remarks")
        loDtaRow.Item("sField17") = p_oDTSrce(lnRow).Item("SMS Stat")
        loDtaRow.Item("sField18") = p_oDTSrce(lnRow).Item("Call Start")
        loDtaRow.Item("sField19") = p_oDTSrce(lnRow).Item("Call End")
        loDtaRow.Item("sField20") = p_oDTSrce(lnRow).Item("Timestamp")
        loDtaRow.Item("sField21") = p_oDTSrce(lnRow).Item("Source No")
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