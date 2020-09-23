VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code by Kailash Nadh , 15 yrs , India
'bnkailash@sify.com , http;//kbn.rom.cd

'This code is not fully mine,
'a little of its part is someone elses!

Option Explicit
Private Declare Function GetDC Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "USER32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private lngDC As Long
Private blnLoop As Boolean
Dim m1 As Integer, m2 As Integer

Private Sub Form_Click()
blnLoop = vbFalse
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
blnLoop = vbFalse
End Sub

Sub doeffect()
Dim intX As Integer, intY As Integer
Dim intI As Integer, intJ As Integer
Dim intWidth As Integer, intHeight As Integer
intWidth = Screen.Width / Screen.TwipsPerPixelX
intHeight = Screen.Height / Screen.TwipsPerPixelY
frmMain.Width = Screen.Width
frmMain.Height = Screen.Height
lngDC = GetDC(0)
Call BitBlt(hDC, 0, 0, intWidth, intHeight, lngDC, 0, 0, vbSrcCopy)
frmMain.Visible = vbTrue
frmMain.AutoRedraw = vbFalse
Randomize
blnLoop = vbTrue
Do While blnLoop = vbTrue
intX = (intWidth - 128) * Rnd
intY = (intHeight - 128) * Rnd
intI = m1 * Rnd - 1 ' The changes in values of m1 & m2 decides the effect
intJ = m2 * Rnd - 1
Call BitBlt(frmMain.hDC, intX + intI, intY + intJ, 128, 128, frmMain.hDC, intX, intY, vbSrcCopy)
DoEvents
Loop
Set frmMain = Nothing
Call ReleaseDC(0, lngDC)
End
End Sub

Private Sub Form_Load()
On Error GoTo er1
Dim ef As Integer
ef = GetSetting("MeltSCR", "Effect", "Effect")

Select Case ef
Case 0
m1 = 2: m2 = 2

Case 1
m1 = 20: m2 = 20

Case 2
m1 = 9: m2 = 9

Case 3
m1 = 0: m2 = 0

Case 4
m1 = 3: m2 = 3

Case 5
m1 = 5: m2 = 5

Case 6
m1 = 10000: m2 = 10000

Case 7
m1 = 1000: m2 = 1000

Case 8
m1 = 10: m2 = 2

Case 9
m1 = 2: m2 = 10

End Select
doeffect
Exit Sub
er1:
m1 = 2: m2 = 2
doeffect
End Sub
