'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Monitors Quick Match Result. 
'
' Copyright 2020 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  jep [ 01/06/2020 09:52 am ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver

Public Class QMResult
    Private p_oDTDetl As DataTable
    Private p_oDTEvlt As DataTable
    Private p_sBranchCD As String
    Private p_oApp As GRider
    Private p_bIsEvaluatr As Boolean

    Public ReadOnly Property ItemCount() As Integer
        Get
            Return p_oDTDetl.Rows.Count
        End Get
    End Property

    WriteOnly Property isEvaluator() As Boolean
        Set(ByVal Value As Boolean)
            p_bIsEvaluatr = Value
        End Set
    End Property

    'Property Detail(Integer, Integer)
    Public ReadOnly Property Detail(ByVal Row As Integer, ByVal Index As Integer) As Object
        Get
            If Not IsNothing(p_oDTDetl) Then
                Return p_oDTDetl(Row).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get
    End Property

    'Property Detail(Integer, String)
    Public ReadOnly Property Detail(ByVal Row As Integer, ByVal Index As String) As Object
        Get
            If Not IsNothing(p_oDTDetl) Then
                Return p_oDTDetl(Row).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get
    End Property

    Public Sub ExtractRecord()
        Dim lsSQL As String

        lsSQL = Replace(IIf(p_bIsEvaluatr, getSQ_Evaluate, getSQ_Detail), "«01»", p_sBranchCD, 1)

        p_oDTDetl = p_oApp.ExecuteQuery(lsSQL)
    End Sub

    Private Function getSQ_Detail() As String
        Return "SELECT" & _
                    " a.sTransNox" & _
                    ", b.sBranchCd" & _
                    ", b.sBranchNm" & _
                    ", a.sClientNm" & _
                    ", a.sQMatchNo" & _
                    ", a.dTransact" & _
                    ", d.sAreaDesc" & _
                    ", a.dReceived" & _
                " FROM Credit_Online_Application a" & _
                    ", Branch b" & _
                        " LEFT JOIN Branch_Others c ON b.sBranchCd = c.sBranchCd" & _
                        " LEFT JOIN Branch_Area d ON c.sAreaCode = d.sAreaCode" & _
                " WHERE a.sSourceCd = 'APP'" & _
                    " AND a.sBranchCd = b.sBranchCd" & _
                    " AND a.cEvaluatr = '0'" & _
                    " AND a.cTranStat IN ('0')" & _
                    " AND DATEDIFF(SYSDATE(), a.dTransact) <= " & CDbl(p_oApp.getConfiguration("QMResult"))
    End Function

    Private Function getSQ_Evaluate() As String
        Dim lnValue As Double

        If IsDBNull(p_oApp.getConfiguration("QMResult")) Then
            lnValue = 0
        Else
            lnValue = IIf(p_oApp.getConfiguration("QMResult") = "", 0, p_oApp.getConfiguration("QMResult"))
        End If

        Return "SELECT * FROM (" & _
                    "SELECT" & _
                        " a.sTransNox" & _
                        ", b.sBranchCd" & _
                        ", b.sBranchNm" & _
                        ", a.sClientNm" & _
                        ", a.sQMatchNo" & _
                        ", a.dTransact" & _
                        ", d.sAreaDesc" & _
                        ", a.dReceived" & _
                    " FROM Credit_Online_Application a" & _
                        ", Branch b" & _
                            " LEFT JOIN Branch_Others c ON b.sBranchCd = c.sBranchCd" & _
                            " LEFT JOIN Branch_Area d ON c.sAreaCode = d.sAreaCode" & _
                    " WHERE a.sSourceCd = 'APP'" & _
                        " AND a.sBranchCd = b.sBranchCd" & _
                        " AND a.cEvaluatr = '1'" & _
                        " AND a.cTranStat NOT IN('3','4')" & _
                        " AND DATEDIFF( " & dateParm(p_oApp.SysDate) & ", a.dTransact) <= " & lnValue & _
                    " UNION ALL" & _
                    " SELECT" & _
                        " a.sTransNox" & _
                        ", b.sBranchCd" & _
                        ", b.sBranchNm" & _
                        ", a.sClientNm" & _
                        ", a.sQMatchNo" & _
                        ", a.dTransact" & _
                        ", d.sAreaDesc" & _
                        ", a.dReceived" & _
                    " FROM Credit_Online_Application a" & _
                        ", Branch b" & _
                            " LEFT JOIN Branch_Others c ON b.sBranchCd = c.sBranchCd" & _
                            " LEFT JOIN Branch_Area d ON c.sAreaCode = d.sAreaCode" & _
                    " WHERE a.sSourceCd = 'APP'" & _
                        " AND a.sBranchCd = b.sBranchCd" & _
                        " AND a.dTransact = " & dateParm(p_oApp.SysDate) & _
                        " AND a.cTranStat NOT IN('3','4')" & _
                        ")xx GROUP BY sTransNox ORDER BY dReceived"

    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_sBranchCD = foRider.BranchCode
    End Sub
End Class