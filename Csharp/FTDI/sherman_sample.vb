Option Strict On

Imports System.Drawing

Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdRun As System.Windows.Forms.Button
    Friend WithEvents txtStatus As TextBox
    Friend WithEvents cmdStop As Button
    Friend WithEvents tmrComms As Timer
    Friend WithEvents tmrMaster As Timer
    Friend WithEvents txtNumReadings As TextBox
    Friend WithEvents Label11 As Label
    Friend WithEvents txtAngle As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents cmdStatus As Button
    Friend WithEvents picGraphDistance As PictureBox
    Friend WithEvents HScrollDistanceBar As HScrollBar
    Friend WithEvents txtDistBar As TextBox
    Friend WithEvents gpboxDistance As GroupBox
    Friend WithEvents Label10 As Label
    Friend WithEvents lblAvgDistance As Label
    Friend WithEvents Label28 As Label
    Friend WithEvents Label27 As Label
    Friend WithEvents lblMaxDistance As Label
    Friend WithEvents cmdAmpDist4X As Button
    Friend WithEvents gpboxRSSI As GroupBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lblAvgLevel As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents lblMinLevel As Label
    Friend WithEvents lblMaxLevel As Label
    Friend WithEvents picGraphRSSI As PictureBox
    Friend WithEvents cmdAmpRSSI4X As Button
    Friend WithEvents txtRSSIBar As TextBox
    Friend WithEvents HScrollRSSIBar As HScrollBar
    Friend WithEvents cmdSweep As Button
    Friend WithEvents lblMinDistance As Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.cmdRun = New System.Windows.Forms.Button()
        Me.txtStatus = New System.Windows.Forms.TextBox()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.tmrComms = New System.Windows.Forms.Timer(Me.components)
        Me.tmrMaster = New System.Windows.Forms.Timer(Me.components)
        Me.txtNumReadings = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtAngle = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cmdStatus = New System.Windows.Forms.Button()
        Me.picGraphDistance = New System.Windows.Forms.PictureBox()
        Me.HScrollDistanceBar = New System.Windows.Forms.HScrollBar()
        Me.txtDistBar = New System.Windows.Forms.TextBox()
        Me.gpboxDistance = New System.Windows.Forms.GroupBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblAvgDistance = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.lblMaxDistance = New System.Windows.Forms.Label()
        Me.lblMinDistance = New System.Windows.Forms.Label()
        Me.cmdAmpDist4X = New System.Windows.Forms.Button()
        Me.gpboxRSSI = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblAvgLevel = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblMinLevel = New System.Windows.Forms.Label()
        Me.lblMaxLevel = New System.Windows.Forms.Label()
        Me.picGraphRSSI = New System.Windows.Forms.PictureBox()
        Me.cmdAmpRSSI4X = New System.Windows.Forms.Button()
        Me.txtRSSIBar = New System.Windows.Forms.TextBox()
        Me.HScrollRSSIBar = New System.Windows.Forms.HScrollBar()
        Me.cmdSweep = New System.Windows.Forms.Button()
        CType(Me.picGraphDistance, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gpboxDistance.SuspendLayout()
        Me.gpboxRSSI.SuspendLayout()
        CType(Me.picGraphRSSI, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdRun
        '
        Me.cmdRun.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRun.Location = New System.Drawing.Point(25, 17)
        Me.cmdRun.Name = "cmdRun"
        Me.cmdRun.Size = New System.Drawing.Size(113, 56)
        Me.cmdRun.TabIndex = 0
        Me.cmdRun.Text = "RUN"
        '
        'txtStatus
        '
        Me.txtStatus.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtStatus.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatus.Location = New System.Drawing.Point(998, 18)
        Me.txtStatus.Multiline = True
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.ReadOnly = True
        Me.txtStatus.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtStatus.Size = New System.Drawing.Size(894, 55)
        Me.txtStatus.TabIndex = 7
        Me.txtStatus.Visible = False
        '
        'cmdStop
        '
        Me.cmdStop.Enabled = False
        Me.cmdStop.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStop.Location = New System.Drawing.Point(823, 18)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(113, 56)
        Me.cmdStop.TabIndex = 13
        Me.cmdStop.Text = "STOP"
        Me.cmdStop.UseVisualStyleBackColor = True
        '
        'tmrComms
        '
        Me.tmrComms.Interval = 50
        '
        'tmrMaster
        '
        Me.tmrMaster.Interval = 20
        '
        'txtNumReadings
        '
        Me.txtNumReadings.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtNumReadings.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNumReadings.Location = New System.Drawing.Point(454, 31)
        Me.txtNumReadings.Name = "txtNumReadings"
        Me.txtNumReadings.Size = New System.Drawing.Size(47, 26)
        Me.txtNumReadings.TabIndex = 87
        Me.txtNumReadings.Text = "5"
        Me.txtNumReadings.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(287, 34)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(175, 21)
        Me.Label11.TabIndex = 86
        Me.Label11.Text = "Number of Readings"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtAngle
        '
        Me.txtAngle.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtAngle.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAngle.Location = New System.Drawing.Point(227, 31)
        Me.txtAngle.Name = "txtAngle"
        Me.txtAngle.Size = New System.Drawing.Size(47, 26)
        Me.txtAngle.TabIndex = 89
        Me.txtAngle.Text = "0"
        Me.txtAngle.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(163, 35)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(76, 21)
        Me.Label12.TabIndex = 88
        Me.Label12.Text = "Angle"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdStatus
        '
        Me.cmdStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStatus.Location = New System.Drawing.Point(652, 18)
        Me.cmdStatus.Name = "cmdStatus"
        Me.cmdStatus.Size = New System.Drawing.Size(144, 56)
        Me.cmdStatus.TabIndex = 100
        Me.cmdStatus.Text = "Show Status"
        Me.cmdStatus.UseVisualStyleBackColor = True
        '
        'picGraphDistance
        '
        Me.picGraphDistance.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picGraphDistance.Image = Global.D2XXUnit_NET.My.Resources.Resources.Whitebackground
        Me.picGraphDistance.Location = New System.Drawing.Point(24, 311)
        Me.picGraphDistance.Name = "picGraphDistance"
        Me.picGraphDistance.Size = New System.Drawing.Size(912, 708)
        Me.picGraphDistance.TabIndex = 101
        Me.picGraphDistance.TabStop = False
        '
        'HScrollDistanceBar
        '
        Me.HScrollDistanceBar.LargeChange = 100
        Me.HScrollDistanceBar.Location = New System.Drawing.Point(517, 98)
        Me.HScrollDistanceBar.Maximum = 3000
        Me.HScrollDistanceBar.Minimum = 50
        Me.HScrollDistanceBar.Name = "HScrollDistanceBar"
        Me.HScrollDistanceBar.Size = New System.Drawing.Size(212, 32)
        Me.HScrollDistanceBar.TabIndex = 102
        Me.HScrollDistanceBar.Value = 50
        '
        'txtDistBar
        '
        Me.txtDistBar.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtDistBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDistBar.Location = New System.Drawing.Point(396, 98)
        Me.txtDistBar.MaxLength = 5
        Me.txtDistBar.Name = "txtDistBar"
        Me.txtDistBar.Size = New System.Drawing.Size(101, 31)
        Me.txtDistBar.TabIndex = 103
        Me.txtDistBar.Text = "50"
        Me.txtDistBar.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'gpboxDistance
        '
        Me.gpboxDistance.Controls.Add(Me.Label10)
        Me.gpboxDistance.Controls.Add(Me.lblAvgDistance)
        Me.gpboxDistance.Controls.Add(Me.Label28)
        Me.gpboxDistance.Controls.Add(Me.Label27)
        Me.gpboxDistance.Controls.Add(Me.lblMaxDistance)
        Me.gpboxDistance.Controls.Add(Me.lblMinDistance)
        Me.gpboxDistance.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gpboxDistance.Location = New System.Drawing.Point(25, 88)
        Me.gpboxDistance.Name = "gpboxDistance"
        Me.gpboxDistance.Size = New System.Drawing.Size(344, 207)
        Me.gpboxDistance.TabIndex = 106
        Me.gpboxDistance.TabStop = False
        Me.gpboxDistance.Text = "Distance (mm)"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(15, 95)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(99, 25)
        Me.Label10.TabIndex = 99
        Me.Label10.Text = "Average"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblAvgDistance
        '
        Me.lblAvgDistance.BackColor = System.Drawing.SystemColors.Control
        Me.lblAvgDistance.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblAvgDistance.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAvgDistance.ForeColor = System.Drawing.Color.Black
        Me.lblAvgDistance.Location = New System.Drawing.Point(137, 89)
        Me.lblAvgDistance.Name = "lblAvgDistance"
        Me.lblAvgDistance.Size = New System.Drawing.Size(181, 37)
        Me.lblAvgDistance.TabIndex = 98
        Me.lblAvgDistance.Text = "0"
        Me.lblAvgDistance.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.ForeColor = System.Drawing.Color.Green
        Me.Label28.Location = New System.Drawing.Point(15, 46)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(111, 25)
        Me.Label28.TabIndex = 97
        Me.Label28.Text = "Maximum"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.Color.Blue
        Me.Label27.Location = New System.Drawing.Point(15, 146)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(105, 25)
        Me.Label27.TabIndex = 96
        Me.Label27.Text = "Minimum"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblMaxDistance
        '
        Me.lblMaxDistance.BackColor = System.Drawing.SystemColors.Control
        Me.lblMaxDistance.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMaxDistance.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaxDistance.ForeColor = System.Drawing.Color.Black
        Me.lblMaxDistance.Location = New System.Drawing.Point(137, 40)
        Me.lblMaxDistance.Name = "lblMaxDistance"
        Me.lblMaxDistance.Size = New System.Drawing.Size(181, 37)
        Me.lblMaxDistance.TabIndex = 95
        Me.lblMaxDistance.Text = "0"
        Me.lblMaxDistance.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblMinDistance
        '
        Me.lblMinDistance.BackColor = System.Drawing.SystemColors.Control
        Me.lblMinDistance.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMinDistance.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMinDistance.ForeColor = System.Drawing.Color.Black
        Me.lblMinDistance.Location = New System.Drawing.Point(137, 140)
        Me.lblMinDistance.Name = "lblMinDistance"
        Me.lblMinDistance.Size = New System.Drawing.Size(181, 37)
        Me.lblMinDistance.TabIndex = 94
        Me.lblMinDistance.Text = "0"
        Me.lblMinDistance.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdAmpDist4X
        '
        Me.cmdAmpDist4X.Font = New System.Drawing.Font("Microsoft Sans Serif", 48.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAmpDist4X.Location = New System.Drawing.Point(396, 153)
        Me.cmdAmpDist4X.Name = "cmdAmpDist4X"
        Me.cmdAmpDist4X.Size = New System.Drawing.Size(132, 112)
        Me.cmdAmpDist4X.TabIndex = 107
        Me.cmdAmpDist4X.Text = "4X"
        Me.cmdAmpDist4X.UseVisualStyleBackColor = True
        '
        'gpboxRSSI
        '
        Me.gpboxRSSI.Controls.Add(Me.Label1)
        Me.gpboxRSSI.Controls.Add(Me.lblAvgLevel)
        Me.gpboxRSSI.Controls.Add(Me.Label3)
        Me.gpboxRSSI.Controls.Add(Me.Label4)
        Me.gpboxRSSI.Controls.Add(Me.lblMinLevel)
        Me.gpboxRSSI.Controls.Add(Me.lblMaxLevel)
        Me.gpboxRSSI.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gpboxRSSI.Location = New System.Drawing.Point(962, 88)
        Me.gpboxRSSI.Name = "gpboxRSSI"
        Me.gpboxRSSI.Size = New System.Drawing.Size(344, 207)
        Me.gpboxRSSI.TabIndex = 108
        Me.gpboxRSSI.TabStop = False
        Me.gpboxRSSI.Text = "RSSI"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 95)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(99, 25)
        Me.Label1.TabIndex = 99
        Me.Label1.Text = "Average"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblAvgLevel
        '
        Me.lblAvgLevel.BackColor = System.Drawing.SystemColors.Control
        Me.lblAvgLevel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblAvgLevel.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAvgLevel.ForeColor = System.Drawing.Color.Black
        Me.lblAvgLevel.Location = New System.Drawing.Point(137, 89)
        Me.lblAvgLevel.Name = "lblAvgLevel"
        Me.lblAvgLevel.Size = New System.Drawing.Size(181, 37)
        Me.lblAvgLevel.TabIndex = 98
        Me.lblAvgLevel.Text = "0"
        Me.lblAvgLevel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Green
        Me.Label3.Location = New System.Drawing.Point(15, 46)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(111, 25)
        Me.Label3.TabIndex = 97
        Me.Label3.Text = "Maximum"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Blue
        Me.Label4.Location = New System.Drawing.Point(15, 146)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(105, 25)
        Me.Label4.TabIndex = 96
        Me.Label4.Text = "Minimum"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblMinLevel
        '
        Me.lblMinLevel.BackColor = System.Drawing.SystemColors.Control
        Me.lblMinLevel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMinLevel.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMinLevel.ForeColor = System.Drawing.Color.Black
        Me.lblMinLevel.Location = New System.Drawing.Point(137, 140)
        Me.lblMinLevel.Name = "lblMinLevel"
        Me.lblMinLevel.Size = New System.Drawing.Size(181, 37)
        Me.lblMinLevel.TabIndex = 95
        Me.lblMinLevel.Text = "0"
        Me.lblMinLevel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblMaxLevel
        '
        Me.lblMaxLevel.BackColor = System.Drawing.SystemColors.Control
        Me.lblMaxLevel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMaxLevel.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaxLevel.ForeColor = System.Drawing.Color.Black
        Me.lblMaxLevel.Location = New System.Drawing.Point(137, 40)
        Me.lblMaxLevel.Name = "lblMaxLevel"
        Me.lblMaxLevel.Size = New System.Drawing.Size(181, 37)
        Me.lblMaxLevel.TabIndex = 94
        Me.lblMaxLevel.Text = "0"
        Me.lblMaxLevel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'picGraphRSSI
        '
        Me.picGraphRSSI.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picGraphRSSI.Image = Global.D2XXUnit_NET.My.Resources.Resources.Whitebackground
        Me.picGraphRSSI.Location = New System.Drawing.Point(962, 311)
        Me.picGraphRSSI.Name = "picGraphRSSI"
        Me.picGraphRSSI.Size = New System.Drawing.Size(912, 708)
        Me.picGraphRSSI.TabIndex = 109
        Me.picGraphRSSI.TabStop = False
        '
        'cmdAmpRSSI4X
        '
        Me.cmdAmpRSSI4X.Font = New System.Drawing.Font("Microsoft Sans Serif", 48.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAmpRSSI4X.Location = New System.Drawing.Point(1333, 153)
        Me.cmdAmpRSSI4X.Name = "cmdAmpRSSI4X"
        Me.cmdAmpRSSI4X.Size = New System.Drawing.Size(132, 112)
        Me.cmdAmpRSSI4X.TabIndex = 112
        Me.cmdAmpRSSI4X.Text = "4X"
        Me.cmdAmpRSSI4X.UseVisualStyleBackColor = True
        '
        'txtRSSIBar
        '
        Me.txtRSSIBar.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtRSSIBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRSSIBar.Location = New System.Drawing.Point(1333, 98)
        Me.txtRSSIBar.MaxLength = 3
        Me.txtRSSIBar.Name = "txtRSSIBar"
        Me.txtRSSIBar.Size = New System.Drawing.Size(101, 31)
        Me.txtRSSIBar.TabIndex = 111
        Me.txtRSSIBar.Text = "10"
        Me.txtRSSIBar.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'HScrollRSSIBar
        '
        Me.HScrollRSSIBar.Location = New System.Drawing.Point(1454, 98)
        Me.HScrollRSSIBar.Maximum = 200
        Me.HScrollRSSIBar.Minimum = 10
        Me.HScrollRSSIBar.Name = "HScrollRSSIBar"
        Me.HScrollRSSIBar.Size = New System.Drawing.Size(212, 32)
        Me.HScrollRSSIBar.TabIndex = 110
        Me.HScrollRSSIBar.Value = 10
        '
        'cmdSweep
        '
        Me.cmdSweep.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSweep.Location = New System.Drawing.Point(788, 98)
        Me.cmdSweep.Name = "cmdSweep"
        Me.cmdSweep.Size = New System.Drawing.Size(148, 88)
        Me.cmdSweep.TabIndex = 113
        Me.cmdSweep.Text = "START SWEEP"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(16, 38)
        Me.ClientSize = New System.Drawing.Size(1904, 1042)
        Me.Controls.Add(Me.cmdSweep)
        Me.Controls.Add(Me.cmdAmpRSSI4X)
        Me.Controls.Add(Me.txtRSSIBar)
        Me.Controls.Add(Me.HScrollRSSIBar)
        Me.Controls.Add(Me.picGraphRSSI)
        Me.Controls.Add(Me.gpboxRSSI)
        Me.Controls.Add(Me.cmdAmpDist4X)
        Me.Controls.Add(Me.gpboxDistance)
        Me.Controls.Add(Me.txtDistBar)
        Me.Controls.Add(Me.HScrollDistanceBar)
        Me.Controls.Add(Me.picGraphDistance)
        Me.Controls.Add(Me.cmdStatus)
        Me.Controls.Add(Me.txtAngle)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.txtNumReadings)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.cmdRun)
        Me.Controls.Add(Me.txtStatus)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "COLA B Communicator and Frame Analyser"
        CType(Me.picGraphDistance, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gpboxDistance.ResumeLayout(False)
        Me.gpboxDistance.PerformLayout()
        Me.gpboxRSSI.ResumeLayout(False)
        Me.gpboxRSSI.PerformLayout()
        CType(Me.picGraphRSSI, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    ' Reasons for timeout
    ' 1. Last byte timeout - Done

    ' Reasons for retry
    ' 1. Failure to meet response length requirements - Done
    ' 2. Failure to pass CRC - No retry - stop comms.....for now
    ' 3. Failure to receive ACK in response - No retry - stop comms.....for now
    ' 4. Faiure to  meet response data requirements - Done

#Region "Constants"

    Public Const DEVICE_ID As String = "USB to IO-Link 1"
    Public Const MAX_RETRIES As Integer = 6
    Public Const RETRY_DELAY As Integer = 100                                          ' In milliseconds
    Public Const STANDARD_TMRCOMM_INTERVAL As Integer = 10          ' In milliseconds
    Public Const LAST_BYTE_COMMS_TIMEOUT As Integer = 10               ' In milliseconds

#Region "COLA B"
    'COLA B
    Public Const COLA_B_STX_LENGTH As Integer = 4
    Public Const COLA_B_LENGTH_LENGTH As Integer = 4
    Public Const COLA_B_CRC_LENGTH As Integer = 1

    Public Const COLA_B_MIN_DIST_OFFSET As Integer = 21
    Public Const COLA_B_MIN_DIST_LENGTH As Integer = 4
    Public Const COLA_B_MAX_DIST_OFFSET As Integer = 25
    Public Const COLA_B_MAX_DIST_LENGTH As Integer = 4
    Public Const COLA_B_DIST_SUM_OFFSET As Integer = 29
    Public Const COLA_B_DIST_SUM_LENGTH As Integer = 4
    Public Const COLA_B_DIST_SQ_SUM_OFFSET As Integer = 33
    Public Const COLA_B_DIST_SQ_SUM_LENGTH As Integer = 4
    Public Const COLA_B_DIST_SQ_SUM_OVERFLOW_OFFSET As Integer = 37
    Public Const COLA_B_DIST_SQ_SUM_OVERFLOW_LENGTH As Integer = 4

    Public Const COLA_B_MIN_LEVEL_OFFSET As Integer = 41
    Public Const COLA_B_MIN_LEVEL_LENGTH As Integer = 4
    Public Const COLA_B_MAX_LEVEL_OFFSET As Integer = 45
    Public Const COLA_B_MAX_LEVEL_LENGTH As Integer = 4
    Public Const COLA_B_LEVEL_SUM_OFFSET As Integer = 49
    Public Const COLA_B_LEVEL_SUM_LENGTH As Integer = 4
    Public Const COLA_B_LEVEL_SUM_SQ_OFFSET As Integer = 53
    Public Const COLA_B_LEVEL_SUM_SQ_LENGTH As Integer = 4
    Public Const COLA_B_LEVEL_SUM_SQ_OVERFLOW_OFFSET As Integer = 57
    Public Const COLA_B_LEVEL_SUM_SQ_OVERFLOW_LENGTH As Integer = 4

    Public Const COLA_B_ACK As Byte = &H41                                      ' 'A'
    Public Const COLA_B_ACK_POS_BY_INDEX As Integer = COLA_B_STX_LENGTH + COLA_B_LENGTH_LENGTH + 1
    Public Const COLA_B_ACK_POS_BY_NAME As Integer = COLA_B_STX_LENGTH + COLA_B_LENGTH_LENGTH + 2

    Private Const COLA_B_STX As String = "02020202"

    Public Const COLA_B_COMMAND_METHOD_POS As Integer = COLA_B_STX_LENGTH + COLA_B_LENGTH_LENGTH + 2
    Private Const COLA_B_BY_INDEX As String = "I"
    Private Const COLA_B_BY_NAME As String = "N"

#End Region

#Region "States"

    Public Const TMR_MASTER_STATE_LOGIN As Integer = 0
    Public Const TMR_MASTER_STATE_MCAVERAGINGDEPTH As Integer = 1
    Public Const TMR_MASTER_STATE_CLEAR_STATISTICOUTDONE_FLAG As Integer = 2
    Public Const TMR_MASTER_STATE_READ_STATISTICOUTDONE_FLAG As Integer = 3
    Public Const TMR_MASTER_STATE_MCDISTRSSISTATISTIC As Integer = 4

    Public Const TMR_MASTER_SUB_STATE_DISPATCH As Integer = 0
    Public Const TMR_MASTER_SUB_STATE_IN_PROGRESS As Integer = 1
    Public Const TMR_MASTER_SUB_STATE_JUST_COMPLETED As Integer = 2

    Public Const TMR_COMMS_STATE_SEND As Integer = 0
    Public Const TMR_COMMS_STATE_GET_RESPONSE As Integer = 1
    Public Const TMR_COMMS_STATE_CHECK_REPONSE As Integer = 2

#End Region

#Region "Commands"

    Public Const CMD_PRODUCTION_LOGIN As String = "734D49000006E80C744D"
    Public Const RES_PRODUCTION_LOGIN_LENGTH As Integer = 15


    Public Const CMD_MCAVERAGINGDEPTH As String = "73574E204D43414420"
    Public Const RES_MCAVERAGINGDEPTH_LENGTH As Integer = 18

    Public Const CMD_CLEAR_STATISTICOUTDONE_FLAG As String = "73574E204D4348532000"
    Public Const RES_CLEAR_STATISTICOUTDONE_FLAG_LENGTH As Integer = 18

    Public Const CMD_READ__STATISTICOUTDONE_FLAG As String = "73524E204D434853"
    Public Const RES_READ__STATISTICOUTDONE_FLAG_LENGTH As Integer = 19

    Public Const CMD_MCDISTRSSISTATISTIC As String = "73524E204D435354"
    Public Const RES_MCDISTRSSISTATISTIC_LENGTH As Integer = 62

#End Region

    Public Const OK As Boolean = False
    Public Const NOT_OK As Boolean = True

    Public Const STATISTICOUTDONE As Byte = &H1

    Public Const GRAPH_DATA_GAP As Integer = 40                               ' In pixels

    Public Const GRAPH_DISTANCE_V_MAGNIFY_STANDARD As Integer = 10

    Public Const SWEEP_START As Integer = -100
    Public Const SWEEP_END As Integer = 100

#End Region

#Region "Variables"

    Dim tmrCommsState As Integer
    Dim tmrCommsResult As Boolean
    Dim stringToSend As String
    Dim expectedResponseLength As Integer

    Dim retryCount As Integer
    Dim commsTimeoutCount As Integer

    Dim commStateMachineDone As Boolean

    Dim tmrMasterState As Integer
    Dim masterSubState As Integer
    Dim stopFlag As Boolean

    Dim receiveBuffer As String
    Dim receivedBufferInByteArray(200) As Byte

    Dim FTStatus As Integer
    Dim FTHandle As Integer

    Dim MCAvgDepthCmd As String
    Dim numReadings As Integer

    Dim oldMinimumPoint As Integer
    Dim newMinimumPoint As Integer

    Dim oldAveragePoint As Integer
    Dim newAveragePoint As Integer

    Dim oldMaximumPoint As Integer
    Dim newMaximumPoint As Integer

    Dim oldRSSIMinimumPoint As Integer
    Dim newRSSIMinimumPoint As Integer

    Dim oldRSSIAveragePoint As Integer
    Dim newRSSIAveragePoint As Integer

    Dim oldRSSIMaximumPoint As Integer
    Dim newRSSIMaximumPoint As Integer

    Dim AmpDist4X As Boolean = False
    Dim ampDistFactor As Integer = 1

    Dim AmpRSSI4X As Boolean = False
    Dim ampRSSIFactor As Integer = 1

    Dim doSweep As Boolean = False

    Dim currentSweepAngle As Integer

    Dim angleMeasured As Boolean = False

#End Region

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tmrComms.Interval = STANDARD_TMRCOMM_INTERVAL
        tmrMaster.Interval = STANDARD_TMRCOMM_INTERVAL
    End Sub

    Private Sub cmdRun_Click(sender As Object, e As EventArgs) Handles cmdRun.Click

        txtStatus.Text = ""

        If setupFTDIChip(DEVICE_ID) = OK Then

            oldMinimumPoint = 0
            oldAveragePoint = 0
            oldMaximumPoint = 0

            oldRSSIMinimumPoint = 0
            oldRSSIAveragePoint = 0
            oldRSSIMaximumPoint = 0

            commStateMachineDone = True
            stopFlag = False
            tmrMasterState = TMR_MASTER_STATE_LOGIN
            masterSubState = TMR_MASTER_SUB_STATE_DISPATCH

            MCAvgDepthCmd = CMD_MCAVERAGINGDEPTH + IntegerStringToHexStringWithPreceedingZeroes((CInt(txtAngle.Text) * 100).ToString, 4) _
                                + IntegerStringToHexStringWithPreceedingZeroes(txtNumReadings.Text, 4)

            Select Case CInt(txtAngle.Text)

                Case -90
                    HScrollDistanceBar.Value = 50

                Case -4
                    HScrollDistanceBar.Value = 1500

                Case 4
                    HScrollDistanceBar.Value = 1500

                Case 13
                    HScrollDistanceBar.Value = 2900

                Case 18
                    HScrollDistanceBar.Value = 2900

                Case 70
                    HScrollDistanceBar.Value = 700

                Case 90
                    HScrollDistanceBar.Value = 1000

            End Select
            txtDistBar.Text = HScrollDistanceBar.Value.ToString
            txtRSSIBar.Text = HScrollRSSIBar.Value.ToString

            numReadings = CInt(txtNumReadings.Text)
            txtAngle.Enabled = False
            txtNumReadings.Enabled = False
            cmdRun.Enabled = False
            cmdStop.Enabled = True
            tmrMaster.Enabled = True
        End If
    End Sub

    Function IntegerStringToHexStringWithPreceedingZeroes(ByVal intString As String, ByVal length As Integer) As String
        Dim tempStr As String = ""

        tempStr = tempStr.PadLeft(length - 1, "0"c)
        tempStr = tempStr + Hex(CInt(intString))
        IntegerStringToHexStringWithPreceedingZeroes = tempStr.Substring(Len(tempStr) - length, length)

    End Function

    Function createCOLABFrameString(originalString As String) As String
        Dim i As Integer
        Dim CRC As Byte
        Dim dataLength As Integer
        Dim outstring As String = ""
        Dim byteString(2) As Byte

        dataLength = CInt(Len(originalString) / 2)

        ReDim byteString(dataLength + COLA_B_STX_LENGTH + COLA_B_LENGTH_LENGTH + COLA_B_CRC_LENGTH - 1)

        ' Convert COLA B STX to Hex in an array of bytes
        HexStringToByteArray(byteString, COLA_B_STX, 0)

        byteString(4) = CByte(dataLength / (256 ^ 3))
        byteString(5) = CByte((dataLength - (byteString(7) * 256 ^ 3)) / (256 ^ 2))
        byteString(6) = CByte((dataLength - (byteString(7) * 256 ^ 3) - (byteString(6) * 256 ^ 2)) / 256)
        byteString(7) = CByte(dataLength - (byteString(7) * 256 ^ 3) - (byteString(6) * 256 ^ 2) - (byteString(5) * 256))

        ' Convert from printable string to Hex in an array of bytes
        HexStringToByteArray(byteString, originalString, COLA_B_STX_LENGTH + COLA_B_LENGTH_LENGTH)

        ' Compute CRC
        CRC = CByte(computeCRC(byteString, COLA_B_STX_LENGTH + COLA_B_LENGTH_LENGTH, dataLength))

        ' Add CRC to complete the COLA B frame
        byteString(dataLength + COLA_B_STX_LENGTH + COLA_B_LENGTH_LENGTH) = CRC

        ' Convert byte array back into a string for transmission
        For i = 0 To byteString.Length - 1
            outstring = outstring + Chr(byteString(i))
        Next

        createCOLABFrameString = outstring
    End Function

    Function computeCRC(ByVal byteArray() As Byte, ByVal offset As Integer, ByVal length As Integer) As Integer
        Dim CRC As Integer
        Dim i As Integer

        CRC = byteArray(offset)
        For i = 1 To length - 1
            CRC = CRC Xor byteArray(i + offset)
        Next i

        computeCRC = CRC
    End Function

    Private Sub HexStringToByteArray(ByRef byteArray() As Byte, ByVal originalString As String, ByVal byteArrayOffset As Integer)
        Dim i As Integer
        Dim tempStr As String

        For i = 0 To CInt(Len(originalString) / 2) - 1
            tempStr = originalString(i * 2) + originalString(i * 2 + 1)
            byteArray(i + byteArrayOffset) = CByte(Convert.ToInt16(tempStr, 16))
        Next i
    End Sub

    Function setupFTDIChip(ByVal deviceID As String) As Boolean
        Dim result As Boolean = OK

        FT_Status = FT_GetNumberOfDevices(FT_Device_Count, vbNullChar, FT_LIST_NUMBER_ONLY)
        If FT_Status <> FT_OK Then
            txtStatus.AppendText(vbCrLf + "Failed to query number of FTDI devices.")
            result = NOT_OK
        Else
            txtStatus.AppendText(vbCrLf + "Successfully queried number of FTDI devices.")
        End If

        If FT_Device_Count > 0 Then

            txtStatus.AppendText(vbCrLf + "Number of devices detected = " + FT_Device_Count.ToString)

            'Open device by device ID
            FTStatus = FT_OpenByDescription(deviceID, FT_OPEN_BY_DESCRIPTION, FTHandle)
            If FTStatus <> FT_OK Then
                txtStatus.AppendText(vbCrLf + DEVICE_ID + " failed to open.")
                result = NOT_OK
            Else
                txtStatus.AppendText(vbCrLf + DEVICE_ID + " opened successfully.")
            End If

            'Reset device
            FTStatus = FT_ResetDevice(FTHandle)
            If FTStatus <> FT_OK Then
                txtStatus.AppendText(vbCrLf + "Failed to reset device.")
                result = NOT_OK
            Else
                txtStatus.AppendText(vbCrLf + "Successfully reset device.")
            End If


            ' Purge buffers
            FTStatus = FT_Purge(FTHandle, FT_PURGE_RX Or FT_PURGE_TX)
            If FTStatus <> FT_OK Then
                txtStatus.AppendText(vbCrLf + "Failed to purge buffers.")
                result = NOT_OK
            Else
                txtStatus.AppendText(vbCrLf + "Buffers purged.")
            End If


            ' Set Baud Rate
            FTStatus = FT_SetBaudRate(FTHandle, FT_BAUD_230400)
            If FTStatus <> FT_OK Then
                txtStatus.AppendText(vbCrLf + "Failed to set baud rate.")
                result = NOT_OK
            Else
                txtStatus.AppendText(vbCrLf + "Baud rate set to" + Str(FT_BAUD_230400) + " bps.")
            End If


            ' Set parameters
            FTStatus = FT_SetDataCharacteristics(FTHandle, FT_DATA_BITS_8, FT_STOP_BITS_1, FT_PARITY_EVEN)
            If FTStatus <> FT_OK Then
                txtStatus.AppendText(vbCrLf + "Failed to set comm data parameters.")
                result = NOT_OK
            Else
                txtStatus.AppendText(vbCrLf + "Successfully set comm parameters.")
            End If


            ' Set Flow Control
            FTStatus = FT_SetFlowControl(FTHandle, FT_FLOW_NONE, FT_FLOW_NONE, FT_FLOW_NONE)
            If FTStatus <> FT_OK Then
                txtStatus.AppendText(vbCrLf + "Failed to configure flow control parameters.")
                result = NOT_OK
            Else
                txtStatus.AppendText(vbCrLf + "Successfully set flow control parameters.")
            End If
        Else
            txtStatus.AppendText(vbCrLf + "No FTDI devices detected.")
            result = NOT_OK
        End If

        setupFTDIChip = result
    End Function

    Private Sub cmdStop_Click(sender As Object, e As EventArgs) Handles cmdStop.Click
        setStopFlag()
    End Sub

    Private Sub setStopFlag()
        stopFlag = True
        txtStatus.AppendText(vbCrLf + "Stopping .............")
    End Sub

    ' This timer acts as a state machine to send out a command and ensures that the length of the expected response is correct.
    Private Sub tmrComms_Tick(sender As Object, e As EventArgs) Handles tmrComms.Tick
        ' Restore interval in case a retry has occured
        tmrComms.Interval = STANDARD_TMRCOMM_INTERVAL

        Select Case (tmrCommsState)

            Case TMR_COMMS_STATE_SEND
                If clearDeviceBuffers() = OK Then

                    ' Send command to device. Will return TRUE if there are problems
                    If sendCommand(stringToSend) = OK Then
                        commsTimeoutCount = 0

                        ' Clear buffer
                        receiveBuffer = ""
                        tmrCommsState = TMR_COMMS_STATE_GET_RESPONSE
                    Else
                        tmrComms.Enabled = False
                        commStateMachineDone = True
                        tmrCommsResult = NOT_OK
                    End If
                Else
                    tmrComms.Enabled = False
                    commStateMachineDone = True
                    tmrCommsResult = NOT_OK
                End If

            Case TMR_COMMS_STATE_GET_RESPONSE
                ' Two ways for getResponseCompleted to return True:
                ' 1. When there is a FTDI issue
                ' 2. When a last byte timeout occurs
                If getResponseCompleted() = True Then
                    If Len(receiveBuffer) = expectedResponseLength * 2 Then
                        'txtStatus.AppendText(vbCrLf + "Received expected response length.")
                        tmrComms.Enabled = False
                        commStateMachineDone = True
                        tmrCommsResult = OK
                    Else
                        txtStatus.AppendText(vbCrLf + "Wrong length of data received.")

                        If OKToRetry() = OK Then
                            tmrComms.Interval = RETRY_DELAY
                            tmrCommsState = TMR_COMMS_STATE_SEND
                        Else
                            tmrComms.Enabled = False
                            txtStatus.AppendText(vbCrLf + "Maximum retries reached.")
                            commStateMachineDone = True
                            tmrCommsResult = NOT_OK
                        End If
                    End If
                End If
        End Select
    End Sub

    Function clearDeviceBuffers() As Boolean
        Dim problemFlag As Boolean

        ' Clear device buffers
        FTStatus = FT_Purge(FTHandle, FT_PURGE_RX Or FT_PURGE_TX)
        If FTStatus <> FT_OK Then
            txtStatus.AppendText(vbCrLf + "Failed to purge buffers.")
            problemFlag = NOT_OK
        Else
            problemFlag = OK
        End If

        clearDeviceBuffers = problemFlag
    End Function

    Function sendCommand(ByVal cmd As String) As Boolean
        Dim problemFlag As Boolean
        Dim bytesWritten As Integer

        FTStatus = FT_Write_String(FTHandle, cmd, CInt(Len(cmd)), bytesWritten)
        If FTStatus <> FT_OK Then
            txtStatus.AppendText(vbCrLf + "Failed to send command.")
            problemFlag = NOT_OK
        Else
            'txtStatus.AppendText(vbCrLf + "Successfully sent out command.")
            problemFlag = OK
        End If

        sendCommand = problemFlag
    End Function

    Function getResponseCompleted() As Boolean
        Dim bytesRead As Integer
        Dim FTRXQBytes As Integer
        Dim tempString As String = ""
        Dim done As Boolean = False

        Dim i As Integer
        Dim tempInt As Byte

        ' Check if there are any bytes to read
        FTStatus = FT_GetQueueStatus(FTHandle, FTRXQBytes)
        If FTStatus <> FT_OK Then
            txtStatus.AppendText(vbCrLf + "Failed to read incoming queue size.")
            done = True
        Else ' No error in reading queue
            If FTRXQBytes > 0 Then
                tempString = Space(FTRXQBytes)
                FTStatus = FT_Read_String(FTHandle, tempString, FTRXQBytes, bytesRead)

                If FTStatus <> FT_OK Then
                    txtStatus.AppendText(vbCrLf + "Failed to read incoming data.")
                    done = True
                Else ' If there are bytes available in the queue
                    commsTimeoutCount = 0

                    For i = 0 To bytesRead - 1
                        tempInt = CByte(Asc(tempString(i)))
                        receiveBuffer = receiveBuffer + tempInt.ToString("X02")
                    Next i
                End If
            Else ' No bytes in queue
                If commsTimeoutCount * STANDARD_TMRCOMM_INTERVAL >= LAST_BYTE_COMMS_TIMEOUT Then
                    commsTimeoutCount = 0
                    done = True
                Else
                    commsTimeoutCount = commsTimeoutCount + 1
                End If
            End If
        End If
        getResponseCompleted = done
    End Function

    Function OKToRetry() As Boolean
        Dim retryState As Boolean

        If retryCount = MAX_RETRIES Then
            retryCount = 0
            retryState = NOT_OK
        Else ' If max retryCount has not been reached
            retryCount = retryCount + 1
            txtStatus.AppendText(vbCrLf + "Incorrect or no response. Retry " + retryCount.ToString + ".")
            txtStatus.AppendText(vbCrLf + "Received response : " + receiveBuffer)
            retryState = OK
        End If
        OKToRetry = retryState
    End Function

    ' This timer acts as a state machine to dispatch commands to tmrComms, performs error checks on the responses and determines the next action.
    Private Sub tmrMaster_Tick(sender As Object, e As EventArgs) Handles tmrMaster.Tick

        Select Case (tmrMasterState)
            Case TMR_MASTER_STATE_LOGIN
                Select Case (masterSubState)
                    Case TMR_MASTER_SUB_STATE_DISPATCH
                        txtStatus.AppendText(vbCrLf + "Attempting to login at production level...........")
                        masterDispatch(CMD_PRODUCTION_LOGIN, RES_PRODUCTION_LOGIN_LENGTH)

                    Case TMR_MASTER_SUB_STATE_IN_PROGRESS
                        checkForProgress()

                    Case TMR_MASTER_SUB_STATE_JUST_COMPLETED
                        If analyseResults(TMR_MASTER_STATE_MCAVERAGINGDEPTH) = NOT_OK Then
                            setStopFlag()
                            txtStatus.AppendText(vbCrLf + "Login failed !")
                        Else
                            txtStatus.AppendText(vbCrLf + "Login successful.")
                        End If
                        checkStopFlag()
                End Select

            Case TMR_MASTER_STATE_MCAVERAGINGDEPTH
                Select Case (masterSubState)
                    Case TMR_MASTER_SUB_STATE_DISPATCH
                        ''txtStatus.AppendText(vbCrLf + "Sending command MCAVERAGINGDEPTH......")
                        masterDispatch(MCAvgDepthCmd, RES_MCAVERAGINGDEPTH_LENGTH)

                    Case TMR_MASTER_SUB_STATE_IN_PROGRESS
                        checkForProgress()

                    Case TMR_MASTER_SUB_STATE_JUST_COMPLETED
                        If analyseResults(TMR_MASTER_STATE_CLEAR_STATISTICOUTDONE_FLAG) = NOT_OK Then
                            setStopFlag()
                            txtStatus.AppendText(vbCrLf + "Failed to get device to start scanning.")
                        End If
                        checkStopFlag()
                End Select

            Case TMR_MASTER_STATE_CLEAR_STATISTICOUTDONE_FLAG
                Select Case (masterSubState)
                    Case TMR_MASTER_SUB_STATE_DISPATCH
                        'txtStatus.AppendText(vbCrLf + "Sucess. Device has starting scanning.......")
                        ''txtStatus.AppendText(vbCrLf + "Sending command to clear done flag......")
                        masterDispatch(CMD_CLEAR_STATISTICOUTDONE_FLAG, RES_CLEAR_STATISTICOUTDONE_FLAG_LENGTH)

                    Case TMR_MASTER_SUB_STATE_IN_PROGRESS
                        checkForProgress()

                    Case TMR_MASTER_SUB_STATE_JUST_COMPLETED
                        If analyseResults(TMR_MASTER_STATE_READ_STATISTICOUTDONE_FLAG) = NOT_OK Then
                            setStopFlag()
                            txtStatus.AppendText(vbCrLf + "Failed to clear done flag in device.")
                        End If
                        checkStopFlag()
                End Select

            Case TMR_MASTER_STATE_READ_STATISTICOUTDONE_FLAG
                Select Case (masterSubState)
                    Case TMR_MASTER_SUB_STATE_DISPATCH
                        ' txtStatus.AppendText(vbCrLf + "Successfully cleared done flag in device.")
                        '' txtStatus.AppendText(vbCrLf + "Reading back done flag......")
                        masterDispatch(CMD_READ__STATISTICOUTDONE_FLAG, RES_READ__STATISTICOUTDONE_FLAG_LENGTH)

                    Case TMR_MASTER_SUB_STATE_IN_PROGRESS
                        checkForProgress()

                    Case TMR_MASTER_SUB_STATE_JUST_COMPLETED
                        If analyseResults(TMR_MASTER_STATE_MCDISTRSSISTATISTIC) = OK Then
                            If receivedBufferInByteArray(17) = &H0 Then
                                tmrMasterState = TMR_MASTER_STATE_READ_STATISTICOUTDONE_FLAG
                                masterSubState = TMR_MASTER_SUB_STATE_DISPATCH
                                '' txtStatus.AppendText(vbCrLf + "Results not ready.....polling device again......")
                            Else
                                ''txtStatus.AppendText(vbCrLf + "Results are ready.")
                            End If
                        Else
                            setStopFlag()
                            txtStatus.AppendText(vbCrLf + "Failed to read back done flag.")
                        End If
                        checkStopFlag()
                End Select

            Case TMR_MASTER_STATE_MCDISTRSSISTATISTIC
                Select Case (masterSubState)
                    Case TMR_MASTER_SUB_STATE_DISPATCH
                        'txtStatus.AppendText(vbCrLf + "Scanning completed.")
                        ''txtStatus.AppendText(vbCrLf + "Retrieving results.........")
                        masterDispatch(CMD_MCDISTRSSISTATISTIC, RES_MCDISTRSSISTATISTIC_LENGTH)

                    Case TMR_MASTER_SUB_STATE_IN_PROGRESS
                        checkForProgress()

                    Case TMR_MASTER_SUB_STATE_JUST_COMPLETED
                        If analyseResults(TMR_MASTER_STATE_MCAVERAGINGDEPTH) = OK Then
                            ''txtStatus.AppendText(vbCrLf + "Successfully retrieved scanning results from device.")
                            If doSweep = True Then
                                updateUI(CInt(picGraphDistance.Width / 200))
                            Else
                                updateUI(GRAPH_DATA_GAP)
                            End If

                            If doSweep = True Then
                                txtStatus.AppendText(vbCrLf + "Angle scanned : " + currentSweepAngle.ToString)
                                currentSweepAngle = currentSweepAngle + 1
                                If currentSweepAngle <= SWEEP_END Then
                                    MCAvgDepthCmd = CMD_MCAVERAGINGDEPTH + IntegerStringToHexStringWithPreceedingZeroes((currentSweepAngle * 100).ToString, 4) _
                                + IntegerStringToHexStringWithPreceedingZeroes(txtNumReadings.Text, 4)
                                Else
                                    doSweep = False
                                    setStopFlag()
                                    cmdSweep.Text = "START SWEEP"
                                    txtStatus.AppendText(vbCrLf + "Sweep from " + SWEEP_START.ToString + " to " + SWEEP_END.ToString + " degrees completed.")
                                End If
                            End If
                        Else
                            setStopFlag()
                            txtStatus.AppendText(vbCrLf + "Failed to retrieve scan results from device.")
                        End If
                        checkStopFlag()
                End Select
        End Select
    End Sub

    Private Sub updateUI(ByVal gap As Integer)

        newMinimumPoint = bytesToIntegerLittleEndian(receivedBufferInByteArray, COLA_B_MIN_DIST_OFFSET, COLA_B_MIN_DIST_LENGTH)
        lblMinDistance.Text = newMinimumPoint.ToString

        newMaximumPoint = bytesToIntegerLittleEndian(receivedBufferInByteArray, COLA_B_MAX_DIST_OFFSET, COLA_B_MAX_DIST_LENGTH)
        lblMaxDistance.Text = newMaximumPoint.ToString

        newAveragePoint = CInt((bytesToIntegerLittleEndian(receivedBufferInByteArray, COLA_B_DIST_SUM_OFFSET, COLA_B_DIST_SUM_LENGTH)) / numReadings)
        lblAvgDistance.Text = newAveragePoint.ToString

        'lblDistSqSum.Text = ((bytesToIntegerLittleEndian(receivedBufferInByteArray, COLA_B_DIST_SQ_SUM_OFFSET, COLA_B_DIST_SQ_SUM_LENGTH))).ToString
        'lblDistSqSumOverflow.Text = ((bytesToIntegerLittleEndian(receivedBufferInByteArray, COLA_B_DIST_SQ_SUM_OVERFLOW_OFFSET, COLA_B_DIST_SQ_SUM_OVERFLOW_LENGTH))).ToString

        newRSSIMinimumPoint = bytesToIntegerLittleEndian(receivedBufferInByteArray, COLA_B_MIN_LEVEL_OFFSET, COLA_B_MIN_LEVEL_LENGTH)
        lblMinLevel.Text = newRSSIMinimumPoint.ToString

        newRSSIMaximumPoint = bytesToIntegerLittleEndian(receivedBufferInByteArray, COLA_B_MAX_LEVEL_OFFSET, COLA_B_MAX_LEVEL_LENGTH)
        lblMaxLevel.Text = newRSSIMaximumPoint.ToString

        newRSSIAveragePoint = CInt((bytesToIntegerLittleEndian(receivedBufferInByteArray, COLA_B_LEVEL_SUM_OFFSET, COLA_B_LEVEL_SUM_LENGTH)) / numReadings)
        lblAvgLevel.Text = newRSSIAveragePoint.ToString

        'lblLevelSqSum.Text = ((bytesToIntegerLittleEndian(receivedBufferInByteArray, COLA_B_LEVEL_SUM_SQ_OFFSET, COLA_B_LEVEL_SUM_SQ_LENGTH))).ToString
        'lblLevelSqSumOverflow.Text = ((bytesToIntegerLittleEndian(receivedBufferInByteArray, COLA_B_LEVEL_SUM_SQ_OVERFLOW_OFFSET, COLA_B_LEVEL_SUM_SQ_OVERFLOW_LENGTH))).ToString

        updateGraphDistance(convertValueToPoint(CInt(picGraphDistance.Height / 2), oldMinimumPoint, HScrollDistanceBar.Value, ampDistFactor),
                                   convertValueToPoint(CInt(picGraphDistance.Height / 2), newMinimumPoint, HScrollDistanceBar.Value, ampDistFactor),
                                    convertValueToPoint(CInt(picGraphDistance.Height / 2), oldAveragePoint, HScrollDistanceBar.Value, ampDistFactor),
                                   convertValueToPoint(CInt(picGraphDistance.Height / 2), newAveragePoint, HScrollDistanceBar.Value, ampDistFactor),
                                convertValueToPoint(CInt(picGraphDistance.Height / 2), oldMaximumPoint, HScrollDistanceBar.Value, ampDistFactor),
                                   convertValueToPoint(CInt(picGraphDistance.Height / 2), newMaximumPoint, HScrollDistanceBar.Value, ampDistFactor), gap)

        updateGraphRSSI(convertValueToPoint(CInt(picGraphRSSI.Height / 2), oldRSSIMinimumPoint, HScrollRSSIBar.Value, ampRSSIFactor),
                                   convertValueToPoint(CInt(picGraphRSSI.Height / 2), newRSSIMinimumPoint, HScrollRSSIBar.Value, ampRSSIFactor),
                                    convertValueToPoint(CInt(picGraphRSSI.Height / 2), oldRSSIAveragePoint, HScrollRSSIBar.Value, ampRSSIFactor),
                                   convertValueToPoint(CInt(picGraphRSSI.Height / 2), newRSSIAveragePoint, HScrollRSSIBar.Value, ampRSSIFactor),
                                convertValueToPoint(CInt(picGraphRSSI.Height / 2), oldRSSIMaximumPoint, HScrollRSSIBar.Value, ampRSSIFactor),
                                   convertValueToPoint(CInt(picGraphRSSI.Height / 2), newRSSIMaximumPoint, HScrollRSSIBar.Value, ampRSSIFactor), gap)


        oldMinimumPoint = newMinimumPoint
        oldAveragePoint = newAveragePoint
        oldMaximumPoint = newMaximumPoint

        oldRSSIMinimumPoint = newRSSIMinimumPoint
        oldRSSIAveragePoint = newRSSIAveragePoint
        oldRSSIMaximumPoint = newRSSIMaximumPoint

    End Sub

    Function convertValueToPoint(ByVal graphCentre As Integer, ByVal value As Integer, ByVal scrollbarValue As Integer, ByVal ampFactor As Integer) As Integer
        convertValueToPoint = graphCentre + (scrollbarValue - value) * GRAPH_DISTANCE_V_MAGNIFY_STANDARD * ampFactor
    End Function

    Private Sub updateGraphDistance(ByVal oldMin As Integer, ByVal newMin As Integer, ByVal oldAvg As Integer, ByVal newAvg As Integer, ByVal oldMax As Integer, ByVal newMax As Integer, ByVal gap As Integer)
        ' Make the Bitmap and Graphics objects

        Dim wid As Integer = picGraphDistance.ClientSize.Width
        Dim hgt As Integer = picGraphDistance.ClientSize.Height
        Dim bm As New Bitmap(wid, hgt)
        Dim gr As Graphics = Graphics.FromImage(bm)

        ' Erase the right edge and draw guide lines.
        For i As Integer = -gap To 0
            gr.DrawLine(Pens.Yellow, wid + i, 0, wid + i, hgt - 1)
        Next i

        gr.DrawImage(picGraphDistance.Image, -gap, 0)           ' Move the old data one pixel to the left.

        ' Draw horizontal guide lines - must use center of graph as reference
        For i As Integer = CInt(picGraphDistance.ClientSize.Height / 2) To picGraphDistance.ClientSize.Height Step (GRAPH_DISTANCE_V_MAGNIFY_STANDARD * ampDistFactor)
            gr.DrawLine(Pens.Blue, wid - gap, i, wid, i)
        Next

        For i As Integer = CInt(picGraphDistance.ClientSize.Height / 2) To 0 Step -GRAPH_DISTANCE_V_MAGNIFY_STANDARD * ampDistFactor
            gr.DrawLine(Pens.Blue, wid - gap, i, wid, i)
        Next

        Dim redPen As New Pen(Brushes.Red)
        redPen.Width = 2.0F

        ' Draw reference line at middle of graph
        gr.DrawLine(redPen, wid - gap, CInt(picGraphDistance.ClientSize.Height / 2), wid, CInt(picGraphDistance.ClientSize.Height / 2))

        ' Draw minimum distance
        Dim myBlueBrush As New System.Drawing.SolidBrush(Color.Blue)

        gr.FillEllipse(myBlueBrush, wid - gap - 5, oldMin - 5, 10, 10)
        Dim bluePen As New Pen(Brushes.Blue)
        bluePen.Width = 3.0F
        gr.DrawLine(bluePen, wid - gap, oldMin, wid, newMin)

        ' Draw average distance
        Dim myBrush As New System.Drawing.SolidBrush(Color.Black)

        gr.FillEllipse(myBrush, wid - gap - 5, oldAvg - 5, 10, 10)
        Dim blackPen As New Pen(Brushes.Black)
        blackPen.Width = 3.0F
        gr.DrawLine(blackPen, wid - gap, oldAvg, wid, newAvg)

        ' Draw maximum distance
        Dim myGreenBrush As New System.Drawing.SolidBrush(Color.Green)

        gr.FillEllipse(myGreenBrush, wid - gap - 5, oldMax - 5, 10, 10)
        Dim greenPen As New Pen(Brushes.Green)
        greenPen.Width = 3.0F
        gr.DrawLine(greenPen, wid - gap, oldMax, wid, newMax)

        ' Display the result.
        picGraphDistance.Image = bm
        picGraphDistance.Refresh()

        gr.Dispose()
    End Sub

    Private Sub updateGraphRSSI(ByVal oldMin As Integer, ByVal newMin As Integer, ByVal oldAvg As Integer, ByVal newAvg As Integer, ByVal oldMax As Integer, ByVal newMax As Integer, ByVal gap As Integer)
        ' Make the Bitmap and Graphics objects

        Dim wid As Integer = picGraphRSSI.ClientSize.Width
        Dim hgt As Integer = picGraphRSSI.ClientSize.Height
        Dim bm As New Bitmap(wid, hgt)
        Dim gr As Graphics = Graphics.FromImage(bm)

        ' Erase the right edge and draw guide lines.
        For i As Integer = -gap To 0
            gr.DrawLine(Pens.Yellow, wid + i, 0, wid + i, hgt - 1)
        Next i

        gr.DrawImage(picGraphRSSI.Image, -gap, 0)           ' Move the old data one pixel to the left.

        ' Draw horizontal guide lines - must use center of graph as reference
        For i As Integer = CInt(picGraphRSSI.ClientSize.Height / 2) To picGraphRSSI.ClientSize.Height Step (GRAPH_DISTANCE_V_MAGNIFY_STANDARD * ampRSSIFactor)
            gr.DrawLine(Pens.Blue, wid - gap, i, wid, i)
        Next

        For i As Integer = CInt(picGraphRSSI.ClientSize.Height / 2) To 0 Step -GRAPH_DISTANCE_V_MAGNIFY_STANDARD * ampRSSIFactor
            gr.DrawLine(Pens.Blue, wid - gap, i, wid, i)
        Next

        Dim redPen As New Pen(Brushes.Red)
        redPen.Width = 2.0F

        ' Draw reference line at middle of graph
        gr.DrawLine(redPen, wid - gap, CInt(picGraphRSSI.ClientSize.Height / 2), wid, CInt(picGraphRSSI.ClientSize.Height / 2))

        ' Draw minimum distance
        Dim myBlueBrush As New System.Drawing.SolidBrush(Color.Blue)

        gr.FillEllipse(myBlueBrush, wid - gap - 5, oldMin - 5, 10, 10)
        Dim bluePen As New Pen(Brushes.Blue)
        bluePen.Width = 3.0F
        gr.DrawLine(bluePen, wid - gap, oldMin, wid, newMin)

        ' Draw average distance
        Dim myBrush As New System.Drawing.SolidBrush(Color.Black)

        gr.FillEllipse(myBrush, wid - gap - 5, oldAvg - 5, 10, 10)
        Dim blackPen As New Pen(Brushes.Black)
        blackPen.Width = 3.0F
        gr.DrawLine(blackPen, wid - gap, oldAvg, wid, newAvg)

        ' Draw maximum distance
        Dim myGreenBrush As New System.Drawing.SolidBrush(Color.Green)

        gr.FillEllipse(myGreenBrush, wid - gap - 5, oldMax - 5, 10, 10)
        Dim greenPen As New Pen(Brushes.Green)
        greenPen.Width = 3.0F
        gr.DrawLine(greenPen, wid - gap, oldMax, wid, newMax)

        ' Display the result.
        picGraphRSSI.Image = bm
        picGraphRSSI.Refresh()

        gr.Dispose()
    End Sub

    Function bytesToIntegerLittleEndian(ByRef array() As Byte, ByVal offset As Integer, ByVal numBytes As Integer) As Integer
        Dim i As Integer
        Dim result As Integer = 0

        For i = 0 To numBytes - 1
            result = result + CInt(array(offset + i) * (256 ^ (numBytes - 1 - i)))
        Next i

        bytesToIntegerLittleEndian = result
    End Function

    Private Sub masterDispatch(ByVal command As String, ByVal responseLength As Integer)
        ' Next sub state
        masterSubState = TMR_MASTER_SUB_STATE_IN_PROGRESS

        ' Preparing tmrComms
        retryCount = 0
        stringToSend = createCOLABFrameString(command)
        expectedResponseLength = responseLength
        commStateMachineDone = False
        tmrCommsState = TMR_COMMS_STATE_SEND
        tmrComms.Enabled = True
    End Sub

    Private Sub checkForProgress()
        If commStateMachineDone = True Then
            masterSubState = TMR_MASTER_SUB_STATE_JUST_COMPLETED
        End If
    End Sub

    Function analyseResults(ByVal nextMasterState As Integer) As Boolean
        Dim result As Boolean = NOT_OK

        If tmrCommsResult = OK Then
            If COLABFrameState() = OK Then
                If CRCCheck() = OK Then
                    If checkForAcknowledgement() = OK Then
                        tmrMasterState = nextMasterState
                        masterSubState = TMR_MASTER_SUB_STATE_DISPATCH
                        result = OK
                    End If
                End If
            End If
        End If

        analyseResults = result
    End Function

    Function COLABFrameState() As Boolean
        Dim result As Boolean = NOT_OK

        If InStr(receiveBuffer, COLA_B_STX, vbTextCompare) = 1 Then
            'txtStatus.AppendText(vbCrLf + "COLA B STX is ok.")
            result = OK
        Else
            txtStatus.AppendText(vbCrLf + "COLA B STX not ok.")
        End If

        COLABFrameState = result
    End Function

    Function CRCCheck() As Boolean
        Dim dataLength As Integer
        Dim CRC As Byte
        Dim response As Boolean = NOT_OK

        ' Convert from ASCII in String to Hex in an array of bytes
        HexStringToByteArray(receivedBufferInByteArray, receiveBuffer, 0)

        ' Calculate data length
        dataLength = bytesToIntegerLittleEndian(receivedBufferInByteArray, COLA_B_STX_LENGTH, COLA_B_LENGTH_LENGTH)

        ' Compute CRC
        CRC = CByte(computeCRC(receivedBufferInByteArray, COLA_B_STX_LENGTH + COLA_B_LENGTH_LENGTH, dataLength))

        ' CRC Check
        If CRC = receivedBufferInByteArray(COLA_B_STX_LENGTH + COLA_B_LENGTH_LENGTH + dataLength) Then
            'txtStatus.AppendText(vbCrLf + "COLA B CRC is ok.")
            response = OK
        Else
            txtStatus.AppendText(vbCrLf + "COLA B CRC failed.")
        End If

        CRCCheck = response
    End Function

    Private Sub checkStopFlag()
        If stopFlag = True Then
            stopFlag = False
            tmrMaster.Enabled = False
            txtStatus.AppendText(vbCrLf + "Stopped.")
            cmdRun.Enabled = True
            cmdStop.Enabled = False
            txtAngle.Enabled = True
            txtNumReadings.Enabled = True
            closeConnection()
        End If
    End Sub

    Private Sub closeConnection()

        FTStatus = FT_Close(FTHandle)
        If FTStatus <> FT_OK Then
            txtStatus.AppendText(vbCrLf + DEVICE_ID + " failed to close.")
        Else
            txtStatus.AppendText(vbCrLf + DEVICE_ID + " closed successfully.")
        End If
    End Sub

    Function checkForAcknowledgement() As Boolean
        Dim posToCheck As Integer
        Dim result As Boolean = NOT_OK

        If stringToSend(COLA_B_COMMAND_METHOD_POS) = COLA_B_BY_INDEX Then
            posToCheck = COLA_B_ACK_POS_BY_INDEX
        Else
            posToCheck = COLA_B_ACK_POS_BY_NAME
        End If

        If receivedBufferInByteArray(posToCheck) = COLA_B_ACK Then            '
            'txtStatus.AppendText(vbCrLf + "COLA B response ACK is ok.")
            result = OK
        Else
            txtStatus.AppendText(vbCrLf + "No COLA B ACK response.")
        End If

        checkForAcknowledgement = result
    End Function

    Private Sub cmdStatus_Click(sender As Object, e As EventArgs) Handles cmdStatus.Click
        If txtStatus.Visible = True Then
            txtStatus.Visible = False
            cmdStatus.Text = "Show Status"
        Else
            txtStatus.Visible = True
            cmdStatus.Text = "Hide Status"
        End If
    End Sub

    Private Sub HScrollAverageDistance_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollDistanceBar.Scroll
        txtDistBar.Text = HScrollDistanceBar.Value.ToString
    End Sub

    Private Sub cmdAmpDist4X_Click(sender As Object, e As EventArgs) Handles cmdAmpDist4X.Click
        If AmpDist4X = False Then
            AmpDist4X = True
            ampDistFactor = 4
            cmdAmpDist4X.BackColor = Color.Lime
        Else
            AmpDist4X = False
            ampDistFactor = 1
            cmdAmpDist4X.BackColor = Control.DefaultBackColor
        End If
    End Sub

    Private Sub txtDistBar_TextChanged(sender As Object, e As EventArgs) Handles txtDistBar.TextChanged

        ' If IsNumeric(txtDistBar.Text) Then
        'If CInt(txtDistBar.Text) < HScrollDistanceBar.Minimum Then
        'txtDistBar.Text = HScrollDistanceBar.Minimum.ToString
        'HScrollDistanceBar.Value = CInt(txtDistBar.Text)
        'HScrollDistanceBar.Update()
        'End If
        'If CInt(txtDistBar.Text) > HScrollDistanceBar.Maximum Then
        'txtDistBar.Text = HScrollDistanceBar.Maximum.ToString
        'HScrollDistanceBar.Value = CInt(txtDistBar.Text)
        'HScrollDistanceBar.Update()
        'End If
        'Else
        'txtDistBar.Text = ""
        'End If


        'If Len(txtDistBar.Text) <= Len(HScrollDistanceBar.Maximum.ToString) Then
        'If Len(txtDistBar.Text) >= Len(HScrollDistanceBar.Minimum.ToString) Then

        'If CInt(txtDistBar.Text) < HScrollDistanceBar.Minimum Then
        'txtDistBar.Text = HScrollDistanceBar.Minimum.ToString
        'HScrollDistanceBar.Value = CInt(txtDistBar.Text)
        'HScrollDistanceBar.Update()
        'ElseIf
        'End If

        'End If
        'End If
        'Else
        'txtDistBar.Text = HScrollDistanceBar.Maximum.ToString
        'End If

        'And CInt(txtDistBar.Text) <= HScrollDistanceBar.Maximum 

    End Sub

    Private Sub HScrollRSSIBar_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollRSSIBar.Scroll
        txtRSSIBar.Text = HScrollRSSIBar.Value.ToString
    End Sub

    Private Sub cmdAmpRSSI4X_Click(sender As Object, e As EventArgs) Handles cmdAmpRSSI4X.Click
        If AmpRSSI4X = False Then
            AmpRSSI4X = True
            ampRSSIFactor = 4
            cmdAmpRSSI4X.BackColor = Color.Lime
        Else
            AmpRSSI4X = False
            ampRSSIFactor = 1
            cmdAmpRSSI4X.BackColor = Control.DefaultBackColor
        End If
    End Sub

    Private Sub cmdSweep_Click(sender As Object, e As EventArgs) Handles cmdSweep.Click
        If doSweep = False Then
            currentSweepAngle = SWEEP_START
            doSweep = True
            cmdSweep.Text = "STOP SWEEP"

            txtStatus.Text = ""

            If setupFTDIChip(DEVICE_ID) = OK Then

                oldMinimumPoint = 0
                oldAveragePoint = 0
                oldMaximumPoint = 0

                oldRSSIMinimumPoint = 0
                oldRSSIAveragePoint = 0
                oldRSSIMaximumPoint = 0

                commStateMachineDone = True
                stopFlag = False
                tmrMasterState = TMR_MASTER_STATE_LOGIN
                masterSubState = TMR_MASTER_SUB_STATE_DISPATCH

                MCAvgDepthCmd = CMD_MCAVERAGINGDEPTH + IntegerStringToHexStringWithPreceedingZeroes((currentSweepAngle * 100).ToString, 4) _
                                + IntegerStringToHexStringWithPreceedingZeroes(txtNumReadings.Text, 4)

                txtDistBar.Text = HScrollDistanceBar.Value.ToString
                txtRSSIBar.Text = HScrollRSSIBar.Value.ToString

                numReadings = CInt(txtNumReadings.Text)
                txtAngle.Enabled = False
                txtNumReadings.Enabled = False
                cmdRun.Enabled = False
                cmdStop.Enabled = False
                tmrMaster.Enabled = True
            End If
        Else
            doSweep = False
            cmdSweep.Text = "START SWEEP"
            stopFlag = True
            checkStopFlag()
        End If
    End Sub

End Class
