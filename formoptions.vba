Option Explicit

Public colorLines As Boolean
Public onlyThisSlide As Boolean
Public Cancelled As Boolean

Private Sub UserForm_Initialize()
    Cancelled = True

    ' Form
    Me.BackColor = vbWhite
    Me.StartUpPosition = 1

    ' Header label (if added lblHeader and lblLine)
    On Error Resume Next
    lblHeader.Left = 0
    lblHeader.Top = 0
    lblHeader.Width = Me.InsideWidth
    lblHeader.Height = 28

    lblLine.Left = 0
    lblLine.Top = 28
    lblLine.Width = Me.InsideWidth
    lblLine.Height = 1
    On Error GoTo 0

    ' Layout constants
    Dim pad As Single: pad = 14
    Dim y As Single: y = 44

    chkOnlyThisSlide.Caption = "Only current slide"
    chkColorLines.Caption = "Color connectors"

    chkOnlyThisSlide.Left = pad
    chkOnlyThisSlide.Top = y
    chkOnlyThisSlide.BackStyle = 0

    y = y + 26
    chkColorLines.Left = pad
    chkColorLines.Top = y
    chkColorLines.BackStyle = 0

    ' Buttons
    Dim btnW As Single: btnW = 90
    Dim btnH As Single: btnH = 28
    Dim btnY As Single: btnY = y + 40

    btnCancel.Width = btnW
    btnCancel.Height = btnH
    btnOK.Width = btnW
    btnOK.Height = btnH

    btnCancel.Left = pad
    btnCancel.Top = btnY

    btnOK.Left = Me.InsideWidth - pad - btnW
    btnOK.Top = btnY

    ' Form size to fit
    Me.Height = btnY + btnH + 44
    Me.Width = 300
End Sub

Private Sub btnOK_Click()
    colorLines = (chkColorLines.Value = True)
    onlyThisSlide = (chkOnlyThisSlide.Value = True)
    Cancelled = False
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Закрытие крестиком = Cancel
    Cancelled = True
End Sub

