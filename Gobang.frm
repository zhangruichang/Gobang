VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   9450
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   11565
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   10560
      Top             =   4920
   End
   Begin VB.Timer Timer2 
      Left            =   9600
      Top             =   4920
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9960
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   8100
      Left            =   480
      Picture         =   "Gobang.frx":0000
      ScaleHeight     =   8040
      ScaleWidth      =   8025
      TabIndex        =   0
      Top             =   360
      Width           =   8085
      Begin VB.Line Line1 
         X1              =   360
         X2              =   7680
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Image Image1 
         Height          =   522
         Index           =   0
         Left            =   99
         Top             =   99
         Visible         =   0   'False
         Width           =   522
      End
   End
   Begin VB.PictureBox WindowsMediaPlayer1 
      Height          =   495
      Left            =   9120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OLE OLE1 
      Class           =   "WPS.Document.6"
      Height          =   495
      Left            =   10800
      OleObjectBlob   =   "Gobang.frx":8970
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   855
      Left            =   9120
      TabIndex        =   4
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   735
      Left            =   9120
      TabIndex        =   3
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   975
      Left            =   9120
      TabIndex        =   2
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   9000
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Menu r 
      Caption         =   "��Ϸ"
      Index           =   0
      Begin VB.Menu run 
         Caption         =   "��ʼ����Ϸ"
         Shortcut        =   ^B
      End
      Begin VB.Menu stop 
         Caption         =   "�˳�"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu r 
      Caption         =   "��Ϸ����"
      Index           =   1
      Begin VB.Menu study 
         Caption         =   "ѧϰ��"
         Shortcut        =   ^S
      End
      Begin VB.Menu tiao 
         Caption         =   "��ս��"
         Shortcut        =   ^C
      End
      Begin VB.Menu ying 
         Caption         =   "Ӧս��"
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu r 
      Caption         =   "����"
      Index           =   2
   End
   Begin VB.Menu r 
      Caption         =   "����"
      Index           =   3
      Begin VB.Menu help 
         Caption         =   "ʹ��ָ��"
         Shortcut        =   ^H
      End
      Begin VB.Menu about 
         Caption         =   "����������"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blk(0 To 224) As Integer 'Ϊ1��ʾ��Ӧ�Ŀؼ����˺���
Dim whi(0 To 224) As Integer 'Ϊ1��ʾ��Ӧ�Ŀؼ����˰���
Dim visited(0 To 224) As Integer '��ʾÿһ���µ����ӵ�λ��

Public Time As String '��ʾʱ���ַ���
Public start As Integer '��ʾѡ�����Ϸ����
Public count1 As Integer '��ʾ����Ĵ���
Public hour1%, minute1%, second1% '��ʾ3��ʱ�Ӷ�Ӧ��ʱ����
Public hour2%, minute2%, second2%
Public hour3%, minute3%, second3%

 

Private Sub about_Click()
    frmAbout.Show (1) '�򿪹��������崰��
End Sub



Private Sub help_Click()
 FileName = "\help.doc" '�򿪰����ĵ�
 
 OLE1.SourceDoc = App.Path + FileName
 OLE1.Action = 1
 OLE1.Action = 7
End Sub

Private Sub run_Click() '�����ʼ����Ϸ
    Label2.Caption = "�ڷ����ߣ�"
    Dim n%
    If (start <> -1) Then
        hour1 = 0
        minute1 = 0
        second1 = 0
        hour2 = 0
        minute2 = 0
        second2 = 0
    For n = 0 To 224
        Image1(n).Visible = True
    Next n
    Timer1.Interval = 1000
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer2.Interval = 1000
'Timer3.Enabled = True
'WindowsMediaPlayer1.URL = "FIFA2012.mp3"
'WindowsMediaPlayer1
    WindowsMediaPlayer1.Visible = False
    End If
End Sub

Private Sub study_Click()
    MsgBox "��ѡ����ѧϰ��,�������˶�ս,����ʱ��", 64, "��ʾ"
    start = 1
End Sub

Private Sub tiao_Click()
    MsgBox "��ѡ������ս��,�������˶�ս,����ʱ��", 64, "��ʾ"
    start = 2
End Sub
Private Sub r_Click(Index As Integer) '����
    If Index = 2 Then
        If count1 = 0 Then
            MsgBox "���Ѿ��ص�����,�㻹��ڰ�!", 48, "����"
        Exit Sub
        End If
        count1 = count1 - 1
        blk(visited(count1)) = 0
        whi(visited(count1)) = 0
        Image1(visited(count1)).Picture = LoadPicture()
    End If
End Sub

Private Sub Timer1_Timer() '˫���ܼ���ʱ��ʱ��
    Dim h$, m$, s$
    h = Str(hour1)
    m = Str(minute1)
    s = Str(second1)
    If (hour1 < 10) Then
        h = "0" & h
    End If
    If (minute1 < 10) Then
        m = "0" & m
    End If
    If (second1 < 10) Then
        s = "0" & s
    End If
    Time = "" & h & ":" & m & ":" & s
    Label1.Caption = "˫���ܼ���ʱ:" + Time
    second1 = second1 + 1
    If (second1 = 60) Then
        minute1 = minute1 + 1
        second1 = 0
    End If
    If (minute1 = 60) Then
        hour1 = hour1 + 1
        minute1 = 0
End If



End Sub

Private Sub Form_Load()

    Label1.Caption = " "
    Label2.Caption = " "
    Label3.Caption = " "
    Label4.Caption = ""
    Timer1.Enabled = False
    Dim n As Integer
    For n = 1 To 224
        Load Image1(n) '���������Ҫ����
        If (Image1(n - 1).Left + Image1(n - 1).Width <= 7407) Then
            Image1(n).Left = Image1(n - 1).Left + Image1(n - 1).Width
            Image1(n).Top = Image1(n - 1).Top
        ElseIf (Image1(n - 1).Top + Image1(n - 1).Height < 7929) Then
            Image1(n).Left = Image1(0).Left
            Image1(n).Top = Image1(n - 1).Top + Image1(n - 1).Height
        Else
        End If
        Image1(n).Visible = False

    Next n
    Timer3.Enabled = False

End Sub
Public Sub Computer() '�˹�����
    Dim i%, j%, x%, loc%, a%, m%
    For i = 0 To 14 '���ȼ���ߵ�4������������Ӧ
        For j = 0 To 11
        x = 15 * i + j
        If (blk(x) + blk(x + 1) + blk(x + 2) + blk(x + 3) = 4) Then
            If (j > 0 And j < 11) Then
                If (whi(x - 1) = 0 And blk(x - 1) = 0) Then
                     loc = x - 1
                ElseIf (whi(x + 4) = 0 And blk(x + 4) = 0) Then
                    loc = x + 4
                Else
                End If
            ElseIf (j = 0) Then
                If (whi(x + 4) = 0 And blk(x + 4) = 0) Then
                    loc = x + 4
                End If
            ElseIf (j = 11) Then
                If (whi(x - 1) = 0 And blk(x - 1) = 0) Then
                    loc = x - 1
                End If
            Else
            End If
        Image1(loc).Picture = LoadPicture("white.icon")
        whi(loc) = 1
        visited(count1) = loc
        Label2.Caption = "�����Ǻ����ߣ�"
        count1 = count1 + 1
        Exit Sub
        End If
        Next j
    Next i

    For i = 0 To 11
        For j = 0 To 14
        x = 15 * i + j
        If (blk(x) + blk(x + 15) + blk(x + 30) + blk(x + 45) = 4) Then
                    If (i > 0 And i < 11) Then
                If (whi(x - 15) = 0 And blk(x - 15) = 0) Then
                     loc = x - 15
                ElseIf (whi(x + 60) = 0 And blk(x + 60) = 0) Then
                    loc = x + 60
                Else
                End If
            ElseIf (i = 0) Then
                If (whi(x + 60) = 0 And blk(x + 60) = 0) Then
                    loc = x + 60
                End If
            ElseIf (i = 11) Then
                If (whi(x - 15) = 0 And blk(x - 15) = 0) Then
                    loc = x - 15
                End If
            Else
            End If
        Image1(loc).Picture = LoadPicture("white.icon")
        whi(loc) = 1
        visited(count1) = loc
        Label2.Caption = "�����Ǻ����ߣ�"
        count1 = count1 + 1
        Exit Sub
        End If
        Next j
    Next i
    
    For i = 0 To 11
        For j = 0 To 11
        x = 15 * i + j
        If (blk(x) + blk(x + 16) + blk(x + 32) + blk(x + 48) = 4) Then
            If (i > 0 And i < 11 And j > 0 And j < 11) Then
                If (whi(x - 16) = 0 And blk(x - 16) = 0) Then
                     loc = x - 16
                ElseIf (whi(x + 64) = 0 And blk(x + 64) = 0) Then
                    loc = x + 64
                Else
                End If
            ElseIf ((i = 0 And j < 11) Or (j = 0 And i < 11)) Then
                If (whi(x + 64) = 0 And blk(x + 64) = 0) Then
                    loc = x + 64
                End If
            ElseIf ((i = 11 And j > 0) Or (j = 11 And i > 0)) Then
                If (whi(x - 16) = 0 And blk(x - 16) = 0) Then
                    loc = x - 16
                End If
            Else
            End If
        Image1(loc).Picture = LoadPicture("white.icon")
        whi(loc) = 1
        visited(count1) = loc
        Label2.Caption = "�����Ǻ����ߣ�"
        count1 = count1 + 1
        Exit Sub
        End If
        Next j
    Next i
    
    For i = 0 To 11
        For j = 3 To 14
        x = 15 * i + j
        If (blk(x) + blk(x + 14) + blk(x + 28) + blk(x + 42) = 4) Then
            If (i > 0 And i < 11 And j > 3 And j < 14) Then
                If (whi(x - 14) = 0 And blk(x - 14) = 0) Then
                     loc = x - 14
                ElseIf (whi(x + 56) = 0 And blk(x + 56) = 0) Then
                    loc = x + 56
                Else
                End If
            ElseIf ((i = 0 And j > 3) Or (j = 14 And i < 11)) Then
                If (whi(x + 56) = 0 And blk(x + 56) = 0) Then
                    loc = x + 56
                End If
            ElseIf ((i = 11 And j < 14) Or (j = 3 And i > 0)) Then
                If (whi(x - 14) = 0 And blk(x - 14) = 0) Then
                    loc = x - 14
                End If
            Else
            End If
        Image1(loc).Picture = LoadPicture("white.icon")
        whi(loc) = 1
        visited(count1) = loc
        Label2.Caption = "�����Ǻ����ߣ�"
        count1 = count1 + 1
        Exit Sub
        End If
        Next j
    Next i
    Call f2(loc)
    Image1(loc).Picture = LoadPicture("white.icon")
    
    whi(loc) = 1
    Label2.Caption = "�����Ǻ����ߣ�"
    visited(count1) = loc
    count1 = count1 + 1
   
End Sub
Public Sub f2(ByRef loc As Integer) '�˹�����Ѱ�����һ��(���л򡣡�)
Dim i%, j%, x%, leftlen%, rightlen%, uplen%, downlen%, leftup%, rightdown%, leftdown%, rightup%, maxlen%, m%, n%, visit(0 To 224)
For i = 0 To 14
    For j = 0 To 14
        m = 0
        n = 0
        x = 15 * i + j
        If (blk(x) = 1 And visit(x) = 0) Then
            visit(x) = 1
            Do While x Mod 15 > 0 And blk(x - 1) = 1 And leftlen < 5
                x = x - 1
                visit(x) = 1
                leftlen = leftlen + 1
            Loop
            If (x Mod 15 > 0 And blk(x - 1) = 0 And whi(x - 1) = 0) Then
                m = x - 1
            End If
            Do While x Mod 15 < 14 And blk(x + 1) = 1 And rightlen < 5
                x = x + 1
                visit(x) = 1
                rightlen = rightlen + 1
            Loop
            If (x Mod 15 < 14 And blk(x + 1) = 0 And whi(x + 1) = 0) Then
                n = x + 1
            End If
            If (maxlen < leftlen + rightlen + 1) Then
                maxlen = leftlen + rightlen + 1
                If (m <> 0) Then
                    loc = m
                ElseIf (n <> 0) Then
                    loc = n
                Else
                End If
            End If
            Do While x \ 15 > 0 And blk(x - 15) = 1 And uplen < 5
                x = x - 15
                visit(x) = 1
                uplen = uplen + 1
            Loop
    If (x \ 15 > 0 And blk(x - 15) = 0 And whi(x - 15) = 0) Then
        m = x - 15
    End If
    Do While x \ 15 < 14 And blk(x + 15) = 1 And downlen < 5
        x = x + 15
        visit(x) = 1
        downlen = downlen + 1
    Loop
    If (x \ 15 < 14 And blk(x + 15) = 0 And whi(x + 15) = 0) Then
        n = x + 15
    End If
    If (maxlen < uplen + downlen = 1) Then
        maxlen = uplen + downlen + 1
        If (m <> 0) Then
            loc = m
        ElseIf (n <> 0) Then
            loc = n
        Else
        End If
    End If
    Do While x \ 15 > 0 And x Mod 15 > 0 And blk(x - 16) = 1 And leftup < 5
        x = x - 16
        visit(x) = 1
        leftup = leftup + 1
    Loop
    If (x \ 15 > 0 And x Mod 15 > 0 And blk(x - 16) = 0 And whi(x - 16) = 0) Then
        m = x - 16
    End If
    Do While x \ 15 < 14 And x Mod 15 < 14 And blk(x + 16) = 1 And rightdown < 5
        x = x + 16
        visit(x) = 1
        rightdown = rightdown + 1
    Loop
    If (x \ 15 < 14 And x Mod 15 < 14 And blk(x + 16) = 0 And whi(x + 16) = 0) Then
        n = x + 16
    End If
    If (maxlen < leftup + rightdown + 1) Then
        maxlen = leftup + rightdown + 1
        If (m <> 0) Then
        loc = m
        ElseIf (n <> 0) Then
        loc = n
        Else
        End If
    End If
    Do While x \ 15 > 0 And x Mod 15 < 14 And blk(x - 14) = 1 And rightup < 5
        x = x - 14
        visit(x) = 1
        rightup = rightup + 1
    Loop
    If (x \ 15 > 0 And x Mod 15 < 14 And blk(x - 14) = 0 And whi(x - 14) = 0) Then
        m = x - 14
    End If
    Do While x \ 15 < 14 And x Mod 15 > 0 And blk(x + 14) = 1 And leftdown < 5
        x = x + 14
        visit(x) = 1
        leftdown = leftdown + 1
    Loop
    If (x \ 15 < 14 And x Mod 15 > 0 And blk(x + 14) = 0 And whi(x + 14) = 0) Then
        n = x + 14
    End If
    If (maxlen < rightup + leftdown + 1) Then
        maxlen = rightup + leftdown + 1
        If (m <> 0) Then
            loc = m
        ElseIf (n <> 0) Then
            loc = n
        Else
        End If
        End If
    End If
Next j
Next i

End Sub
Private Sub player(Index As Integer) '������

    Timer2.Enabled = False
    
    
    Dim i%, j%, x%, exchange%
    If (blk(Index) = 1 Or whi(Index) = 1) Then
        MsgBox "�벻Ҫ���¹��Ӵ����壡", 48, "����"
    Else
    
    Image1(Index).Picture = LoadPicture("black.icon")
    blk(Index) = 1
    visited(count1) = Index
    Label2.Caption = "�����ǰ����ߣ�"
    count1 = count1 + 1
    End If
End Sub
Private Sub exchange() '���ֲ����������
Dim exchange As Boolean, x%
For i = 0 To 14
    For j = 0 To 14
    x = 15 * i + j
    If blk(x) = 1 Then
       
        If (i > 0 And j > 0 And blk(x - 16) = 1) Then
            If (i > 1 And j > 1 And whi(x - 32) = 0 And i < 14 And j < 14 And whi(x + 16) = 0) Then
            exchange = 1
            
            End If
        End If
        If (i < 14 And j < 14 And blk(x + 16) = 1) Then
            If (i > 0 And j > 0 And whi(x - 16) = 0 And i < 13 And j < 13 And whi(x + 32) = 0) Then
            exchange = 1
            
            End If
        End If
        If (i > 0 And j < 14 And blk(x - 14) = 1) Then
            If (i > 1 And j < 13 And whi(x - 28) = 0 And i < 14 And j > 0 And whi(x + 14) = 0) Then
            exchange = 1
            
            End If
        End If
        If (i < 14 And j > 0 And blk(x + 14) = 1) Then
            If (i > 0 And j < 14 And whi(x - 14) = 0 And i < 13 And j > 1 And whi(x + 28) = 0) Then
            exchange = 1
            
            End If
        End If
    End If
    Next j
    Next i
    If (exchange = True) Then
        MsgBox "�׷��������", 48, "ѯ��"
        For i = 0 To 14
            For j = 0 To 14
                x = 15 * i + j
                If (blk(x) = 1) Then
                    Image1(x).Picture = LoadPicture("white.icon")
                    blk(x) = 0
                    whi(x) = 1
                ElseIf (whi(x) = 1) Then
                    Image1(x).Picture = LoadPicture("black.icon")
                    whi(x) = 0
                    blk(x) = 1
                Else
                End If
            Next j
        Next i
    count1 = count1 + 1
    End If
    
End Sub



Private Sub Image1_Click(Index As Integer) '����image�ؼ��ĵ���¼�
If (start = 1 Or start = 2) Then '���˶�ս
    If (blk(Index) = 1 Or whi(Index) = 1) Then
        MsgBox "�벻Ҫ���¹��Ӵ����壡", 48, "����"
    Else


        If (count1 Mod 2 = 0) Then
            Image1(Index).Picture = LoadPicture("black.icon")
            blk(Index) = 1
            visited(count1) = Index
            Label2.Caption = "�����ǰ����ߣ�"
        If (start = 2) Then
            Timer2.Enabled = False
            Timer3.Enabled = True
        End If
        ElseIf count1 Mod 2 = 1 Then
        Image1(Index).Picture = LoadPicture("white.icon")
        whi(Index) = 1
        visited(count1) = Index
        Label2.Caption = "�����Ǻ����ߣ�"
        If (start = 2) Then
            Timer2.Enabled = True
            Timer3.Enabled = False
        End If
    Else
End If
count1 = count1 + 1
End If
Else '�˻���ս
Timer1.Enabled = True

If (count1 Mod 2 = 0) Then
    Call player(Index)
End If

If count1 = 3 Then
    Call exchange
End If
If (count1 Mod 2 = 1) Then
    Call Computer
End If
End If

For i = 0 To 14 '�ж�ʤ��
    For j = 0 To 10
        x = 15 * i + j
        If (blk(x) + blk(x + 1) + blk(x + 2) + blk(x + 3) + blk(x + 4) = 5) Then
            MsgBox "�ڷ�ʤ", 48, "��������"
        ElseIf (whi(x) + whi(x + 1) + whi(x + 2) + whi(x + 3) + whi(x + 4) = 5) Then
            MsgBox "�׷�ʤ", 48, "��������"
        Else
        End If
Next j
Next i

For i = 0 To 10
    For j = 0 To 14
        x = 15 * i + j
        If (blk(x) + blk(x + 15) + blk(x + 30) + blk(x + 45) + blk(x + 60) = 5) Then
            MsgBox "�ڷ�ʤ", 48, "��������"
        ElseIf (whi(x) + whi(x + 15) + whi(x + 30) + whi(x + 45) + whi(x + 60) = 5) Then
            MsgBox "�׷�ʤ", 48, "��������"
        Else
        End If
    Next j
Next i

For i = 0 To 10
    For j = 0 To 10
        x = 15 * i + j
        If (blk(x) + blk(x + 16) + blk(x + 32) + blk(x + 48) + blk(x + 64) = 5) Then
            MsgBox "�ڷ�ʤ", 48, "��������"
        ElseIf (whi(x) + whi(x + 16) + whi(x + 32) + whi(x + 48) + whi(x + 64) = 5) Then
            MsgBox "�׷�ʤ", 48, "��������"
        Else
        End If
    Next j
Next i

For i = 0 To 10
    For j = 4 To 14
        x = 15 * i + j
        If (blk(x) + blk(x + 14) + blk(x + 28) + blk(x + 42) + blk(x + 56) = 5) Then
            MsgBox "�ڷ�ʤ", 48, "��������"
        ElseIf (whi(x) + whi(x + 14) + whi(x + 28) + whi(x + 42) + whi(x + 56) = 5) Then
            MsgBox "�׷�ʤ", 48, "��������"
        Else
        End If
    Next j
Next i
End Sub

Private Sub stop_Click()
End
End Sub


Private Sub Timer2_Timer() '�ڷ���ʱ��ʱ��
    If (start = 2) Then
        Timer2.Enabled = True
    Else
        Timer2.Enabled = False
    End If
    Timer2.Interval = 1000
    Dim h$, m$, s$
    h = Str(hour2)
    m = Str(minute2)
    s = Str(second2)
    If (hour2 < 10) Then
        h = "0" & h
    End If
    If (minute2 < 10) Then
        m = "0" & m
    End If
    If (second2 < 10) Then
        s = "0" & s
    End If
    Time = "" & h & ":" & m & ":" & s
    Label3.Caption = "�ڷ���ʱ:" + Time
    second2 = second2 + 1
    If (second2 = 60) Then
        minute2 = minute2 + 1
        second2 = 0
    End If
    If (minute2 = 60) Then
        hour2 = hour2 + 1
        minute2 = 0
    End If
End Sub

Private Sub Timer3_Timer() '�׷���ʱ��ʱ��
    If (start = 2) Then
        Timer3.Enabled = True
    Else
        Timer3.Enabled = False
    End If
    
    Timer3.Interval = 1000
    Dim h$, m$, s$
    h = Str(hour3)
    m = Str(minute3)
    s = Str(second3)
    
    If (hour3 < 10) Then
        h = "0" & h
    End If
    If (minute3 < 10) Then
        m = "0" & m
    End If
    If (second3 < 10) Then
        s = "0" & s
    End If
    Time = "" & h & ":" & m & ":" & s
    Label4.Caption = "�׷���ʱ:" + Time
    second3 = second3 + 1
    If (second3 = 60) Then
        minute3 = minute3 + 1
        second3 = 0
    End If
    If (minute3 = 60) Then
        hour3 = hour3 + 1
        minute3 = 0
    End If

End Sub

Private Sub ying_Click()
    MsgBox "��ѡ����Ӧս��,�����˻���ս,����ʱ��", 64, "��ʾ"
    start = 3
End Sub
