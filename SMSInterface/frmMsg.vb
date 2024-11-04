Public Class frmMsg
    Dim m_Result As Integer

    Public ReadOnly Property MsgResult
        Get
            Return m_Result
        End Get
    End Property

    Private Sub frmMsg_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = 3 Then
            e.Cancel = True
            m_Result = 0
            Me.Hide()
        End If
    End Sub

    Public Sub New(ByVal Prompt As String, Optional ByVal Buttons As clsGlobal.MsgButtonEnum = clsGlobal.MsgButtonEnum.OKOnly, Optional ByVal Title As String = "")
        InitializeComponent()
        txtMsg.Text = Prompt
        Me.Text = Title
        If Buttons = clsGlobal.MsgButtonEnum.OKOnly Then
            btnYes.Visible = False
            btnCancel.Visible = False
            btnOK.Left = 209
            AcceptButton = btnOK
        ElseIf Buttons = clsGlobal.MsgButtonEnum.OKCancel Then
            btnYes.Visible = False
            AcceptButton = btnOK
            CancelButton = btnCancel
        ElseIf Buttons = clsGlobal.MsgButtonEnum.YesNo Then
            btnCancel.Visible = False
            btnOK.Text = "&No"
            btnYes.Left = 129
            btnOK.Left = 209            
        ElseIf Buttons = clsGlobal.MsgButtonEnum.YesNoCancel Then
            btnOK.Text = "&No"
            CancelButton = btnCancel
        End If

    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If btnYes.Visible = False Then
            m_Result = 1
        Else
            m_Result = 2
        End If
        Me.Hide()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        m_Result = 0
        Me.Hide()
    End Sub

    Private Sub btnYes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnYes.Click
        m_Result = 3
        Me.Hide()
    End Sub

    Private Sub frmMsg_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class