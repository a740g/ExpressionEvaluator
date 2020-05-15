' Example code using the Expression Evaluator
' Copyright (c) Samuel Gomes (Blade), 2001-2003
' mailto: v_2samg@hotmail.com

Friend Class FrmMain
    Inherits Form

#Region "Windows Form Designer generated code"

    Public Sub New()
        MyBase.New()
        If m_vb6FormDefInstance Is Nothing Then
            If m_InitializingDefInstance Then
                m_vb6FormDefInstance = Me
            Else
                Try
                    'For the start-up form, the first instance created is the default instance.
                    If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
                        m_vb6FormDefInstance = Me
                    End If
                Catch
                End Try
            End If
        End If
        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As ToolTip
    Public WithEvents CmdClear As Button
    Public WithEvents ChkInstant As CheckBox
    Public WithEvents CmdDefine As Button
    Public WithEvents CmdEvaluate As Button
    Public WithEvents TxtExpression As TextBox
    Public WithEvents Label1 As Label
    Public WithEvents LblResult As Label

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New ComponentModel.Container()
        Dim resources As ComponentModel.ComponentResourceManager = New ComponentModel.ComponentResourceManager(GetType(FrmMain))
        Me.ToolTip1 = New ToolTip(Me.components)
        Me.CmdClear = New Button()
        Me.ChkInstant = New CheckBox()
        Me.CmdDefine = New Button()
        Me.CmdEvaluate = New Button()
        Me.TxtExpression = New TextBox()
        Me.LblResult = New Label()
        Me.Label1 = New Label()
        Me.SuspendLayout()
        '
        'cmdClear
        '
        Me.CmdClear.BackColor = System.Drawing.SystemColors.Control
        Me.CmdClear.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdClear.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CmdClear.Font = New Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdClear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdClear.Location = New Point(384, 64)
        Me.CmdClear.Name = "cmdClear"
        Me.CmdClear.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdClear.Size = New Size(81, 25)
        Me.CmdClear.TabIndex = 4
        Me.CmdClear.Text = "&Clear"
        Me.ToolTip1.SetToolTip(Me.CmdClear, "Clears everything.")
        Me.CmdClear.UseVisualStyleBackColor = False
        '
        'chkInstant
        '
        Me.ChkInstant.BackColor = System.Drawing.SystemColors.Control
        Me.ChkInstant.Cursor = System.Windows.Forms.Cursors.Default
        Me.ChkInstant.Font = New Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkInstant.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ChkInstant.Location = New Point(384, 8)
        Me.ChkInstant.Name = "chkInstant"
        Me.ChkInstant.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ChkInstant.Size = New Size(80, 21)
        Me.ChkInstant.TabIndex = 2
        Me.ChkInstant.Text = "&Instant"
        Me.ToolTip1.SetToolTip(Me.ChkInstant, "Evaluates expression as it is being entered.")
        Me.ChkInstant.UseVisualStyleBackColor = False
        '
        'cmdDefine
        '
        Me.CmdDefine.BackColor = System.Drawing.SystemColors.Control
        Me.CmdDefine.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdDefine.Font = New Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdDefine.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdDefine.Location = New Point(384, 32)
        Me.CmdDefine.Name = "cmdDefine"
        Me.CmdDefine.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdDefine.Size = New Size(81, 25)
        Me.CmdDefine.TabIndex = 3
        Me.CmdDefine.Text = "&Define"
        Me.ToolTip1.SetToolTip(Me.CmdDefine, "Define a symbol.")
        Me.CmdDefine.UseVisualStyleBackColor = False
        '
        'cmdEvaluate
        '
        Me.CmdEvaluate.BackColor = System.Drawing.SystemColors.Control
        Me.CmdEvaluate.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdEvaluate.Font = New Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdEvaluate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdEvaluate.Location = New Point(384, 96)
        Me.CmdEvaluate.Name = "cmdEvaluate"
        Me.CmdEvaluate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdEvaluate.Size = New Size(81, 25)
        Me.CmdEvaluate.TabIndex = 5
        Me.CmdEvaluate.Text = "&Evaluate"
        Me.ToolTip1.SetToolTip(Me.CmdEvaluate, "Evaluate the expression.")
        Me.CmdEvaluate.UseVisualStyleBackColor = False
        '
        'txtExpression
        '
        Me.TxtExpression.AcceptsReturn = True
        Me.TxtExpression.BackColor = System.Drawing.SystemColors.Window
        Me.TxtExpression.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtExpression.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtExpression.Location = New Point(8, 32)
        Me.TxtExpression.MaxLength = 0
        Me.TxtExpression.Multiline = True
        Me.TxtExpression.Name = "txtExpression"
        Me.TxtExpression.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtExpression.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TxtExpression.Size = New Size(369, 89)
        Me.TxtExpression.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.TxtExpression, "Directly enter/edit the expression here.")
        '
        'lblResult
        '
        Me.LblResult.BackColor = System.Drawing.SystemColors.Control
        Me.LblResult.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LblResult.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblResult.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblResult.Location = New Point(8, 8)
        Me.LblResult.Name = "lblResult"
        Me.LblResult.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblResult.Size = New Size(369, 21)
        Me.LblResult.TabIndex = 0
        Me.LblResult.Text = "0"
        Me.LblResult.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.LblResult, "Result.")
        Me.LblResult.UseMnemonic = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New Point(8, 128)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New Size(359, 14)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Copyright © Samuel Gomes, 2001-2020.     mailto: v_2samg@hotmail.com"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'FrmMain
        '
        Me.AcceptButton = Me.CmdEvaluate
        Me.AllowDrop = True
        Me.AutoScaleBaseSize = New Size(6, 15)
        Me.CancelButton = Me.CmdClear
        Me.ClientSize = New Size(470, 148)
        Me.Controls.Add(Me.CmdClear)
        Me.Controls.Add(Me.ChkInstant)
        Me.Controls.Add(Me.CmdDefine)
        Me.Controls.Add(Me.CmdEvaluate)
        Me.Controls.Add(Me.TxtExpression)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LblResult)
        Me.Font = New Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), Icon)
        Me.MaximizeBox = False
        Me.Name = "FrmMain"
        Me.Opacity = 0.99R
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Expression Evaluator"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "Upgrade Support"

    Private Shared m_vb6FormDefInstance As FrmMain
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As FrmMain
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New FrmMain()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal Value As FrmMain)
            m_vb6FormDefInstance = Value
        End Set
    End Property

#End Region

    Private Sub ChkInstant_CheckStateChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles ChkInstant.CheckStateChanged
        TxtExpression_TextChanged(TxtExpression, New EventArgs())
    End Sub

    Private Sub CmdClear_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles CmdClear.Click
        TxtExpression.Text = ""
        LblResult.Text = ""
        LblResult.Text = CStr(Val(LblResult.Text))
    End Sub

    Private Sub CmdDefine_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles CmdDefine.Click
        Static sSymbol, sValue As String

        sSymbol = InputBox("Enter a symbol name:", , sSymbol)
        If sSymbol = "" Then
            Exit Sub
        End If

        sValue = InputBox("Enter value for " & sSymbol & " (nothing removes):", , sValue)
        If sValue = "" Then
            If EvaluatorSymbolTable.IsDefined(sSymbol) Then
                EvaluatorSymbolTable.Delete(sSymbol)
            End If

            CmdEvaluate_Click(CmdEvaluate, New EventArgs())

            Exit Sub
        End If

        On Error GoTo SymbolError

        If EvaluatorSymbolTable.IsDefined(sSymbol) Then
            EvaluatorSymbolTable.Value(sSymbol) = Val(sValue)
        Else
            EvaluatorSymbolTable.Add(sSymbol, Val(sValue))
        End If

        CmdEvaluate_Click(CmdEvaluate, New EventArgs())

        Exit Sub

SymbolError:
        MsgBox(Err.Description, MsgBoxStyle.Critical)
    End Sub

    Private Sub CmdEvaluate_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles CmdEvaluate.Click
        'Trap errors
        On Error GoTo EvalError

        'Evaluate expression
        LblResult.Text = CStr(Evaluate(TxtExpression.Text))
        Exit Sub

EvalError:
        LblResult.Text = Err.Description
    End Sub

    Private Sub FrmMain_Load(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles MyBase.Load
        ' Define some standard symbols
        EvaluatorSymbolTable.Add("PI", 3.1415926535897931)
        EvaluatorSymbolTable.Add("E", 2.7182818284590451)
        EvaluatorSymbolTable.Add("SQ2", 1.4142135623730952)
        EvaluatorSymbolTable.Add("SQ3", 1.7320508075688772)
        EvaluatorSymbolTable.Add("SQ5", 2.23606797749979)
        EvaluatorSymbolTable.Add("GR", 1.6180339887498949)
        EvaluatorSymbolTable.Add("OC", 0.56714329040978384)
        ' Add more some day :)
    End Sub

    Private Sub TxtExpression_TextChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles TxtExpression.TextChanged
        If ChkInstant.CheckState = System.Windows.Forms.CheckState.Checked Then
            CmdEvaluate_Click(CmdEvaluate, New EventArgs())
        End If
    End Sub

    Private Sub TxtExpression_KeyPress(ByVal eventSender As Object, ByVal eventArgs As KeyPressEventArgs) Handles TxtExpression.KeyPress
        Dim KeyAscii As Short = CShort(Asc(eventArgs.KeyChar))
        If KeyAscii = System.Windows.Forms.Keys.Enter Then
            CmdEvaluate_Click(CmdEvaluate, New EventArgs())
        End If
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
End Class
