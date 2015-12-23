Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsDelivery
    Inherits clsBase
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String
    Private oMatrix As SAPbouiCOM.Matrix

    Public Sub New()
        MyBase.New()
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Delivery Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            'Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            '    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            '    If pVal.ItemUID = "1" And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            '        Dim strMessage As String = String.Empty
                            '        If Not oApplication.Utilities.validate_Accounts(oForm, strMessage) Then
                            '            oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '            BubbleEvent = False
                            '        End If
                            '    End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                initializeControls(oForm)
                                oApplication.Utilities.AddFreightControls(oForm, "ODLN")
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            Select Case BusinessObjectInfo.BeforeAction
                Case True

                Case False
                    Select Case BusinessObjectInfo.EventType
                      
                    End Select
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"

    Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
        Try
            
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
