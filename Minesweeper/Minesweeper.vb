Option Explicit On

Imports System.IO

Public Class Minesweeper
#Region "Disposing"
    Implements IDisposable
    ' To detect redundant calls
    Private disposed As Boolean = False

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposed Then
            If disposing Then
                If Not img Is Nothing Then img.Dispose()
                If Not img Is Nothing Then img = Nothing
                If Not rst Is Nothing Then rst.Dispose()
                If Not rst Is Nothing Then rst = Nothing
                If Not animationTimer Is Nothing Then animationTimer.Dispose()
                If Not animationTimer Is Nothing Then animationTimer = Nothing
                MineStates = Nothing
                pMF = Nothing
                pMinefield.Dispose()
                pReset.Dispose()
                fnAnimations = Nothing
                Tiles = Nothing
                TileH = Nothing
                TileW = Nothing
                W = Nothing
                H = Nothing
                MineCount = Nothing
                blnMouseDownMatch = Nothing
                blnGameOver = Nothing
                blnEnableQuestion = Nothing
                blnFirstMove = Nothing
                ' Free other state (managed objects).
            End If
            ' Free your own state (unmanaged objects).
            ' Set large fields to null.

        End If
        Me.disposed = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to 
    ' correctly implement the disposable pattern.
    Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. 
        ' Put cleanup code in
        ' Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        ' Do not change this code. 
        ' Put cleanup code in
        ' Dispose(ByVal disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub
#End Region
#End Region

    'To-Do
    'Add Color chooser
    'View menu item, for TileW/resetW/etc

#Region "Global Variable Declarations"
    Private Const DefaultResetW As Byte = 26 'Constants
    Private Const DefaultResetH As Byte = 26
    Private Const DefaultSmileyW As Byte = 17
    Private Const DefaultSmileyH As Byte = 17
    Private Const DefaultTileW As Byte = 16
    Private Const DefaultTileH As Byte = 16
    Private Const DefaultNumberW As Byte = 13
    Private Const DefaultNumberH As Byte = 23

    Private Tiles(,,) As Byte
    Private GameType As Byte = 0
    Private MineCount As Short = 0
    Private W As Short = 0
    Private H As Short = 0
    Private TileW As Short = 0
    Private TileH As Short = 0
    Private SmileyW As Byte = 17
    Private SmileyH As Byte = 17
    Private ResetH As Byte = 26
    Private ResetW As Byte = 26
    Private BorderBuffer As Byte = 7
    Private BorderExtension As Byte = 0
    Private iExt As Byte = 0
    Private MineIndex As Byte = 0
    Private FlagIndex As Byte = 0
    Private ButtonUpIndex As Byte = 0
    Private QuestionIndex As Byte = 0
    Private CrossIndex As Byte = 0
    Private ResetIndex As Byte = 0
    Private ResetIndex2 As Byte = 0
    Private TimeTakenCounter As Short
    Private MinesLeftCounter As String
    Private MeLeft As Short = 0
    Private MeTop As Short = 0
    Private MeWidth As Short = 0
    Private MeHeight As Short = 0

    Private blnMouseDownMatch As Boolean = False 'Booleans
    Private blnEnableQuestion As Boolean = True
    Private blnGameOver As Boolean = False
    Private blnVictory As Boolean = False
    Private blnFirstMove As Boolean = False
    Private blnScaleNumbers As Boolean = False
    Private blnSelected As Boolean = False

    Private img As Bitmap 'Image varaibles
    Private rst As Bitmap
    Private MinesLeft As Bitmap
    Private TimeTaken As Bitmap
    Private MineStates As List(Of Bitmap)
    Private Numbers As List(Of Bitmap)
    Private pMinefield As PictureBox
    Private pReset As PictureBox
    Private pMinesLeft As PictureBox
    Private pTimeTaken As PictureBox

    Private animationTimer As Timer
    Private timetakenTimer As Timer
    Private clr As Color
    Private fnContent As String
    Private fnAnimations As String
    Private lstHightlight As New List(Of Point)
    Private MBox As New BMB
    Private SF As New SafeFile

    Private MainMenu As MenuStrip
    Private MenuGame As List(Of MenuItem)
    Private MenuHelp As List(Of MenuItem)
    Private objToolTip As New ToolTip
    Private ToolTipDisplay As Boolean = True
    Private SIndex As Byte = 0
    Private panGame As Panel
    Private panBar As Panel

    Private pMF As PMouse
    Private Structure PMouse
        Public MouseX As Short
        Public MouseY As Short
        Public MouseDownX As Short
        Public MouseDownY As Short
        Public blnLMDown As Boolean
        Public blnRMDown As Boolean
        Public blnMouseOffPicture As Boolean
    End Structure

    Private Structure Animation
        Public FrameCount As Byte
        Public FrameIndex As Short
        Public FrameTimer As List(Of Short)
        Public InfiniteRepeat As Boolean
        Public Repeats As Byte
        Public RepeatCounter As Byte
    End Structure

#End Region

#Region "Load Minesweeper"
    Public Sub New(ByVal prnt As Control)
        FindResources()
        NewGame(prnt, 9, 9, 10, 0, 16, 16, Nothing, False)
    End Sub

    Public Sub New(ByVal prnt As Control, ByRef mw As Short, ByRef mh As Short, ByRef m As Short, Optional ByRef gtype As Byte = 0, Optional ByRef tw As Short = 16, Optional ByRef th As Short = 16, Optional ByRef tClr As Color = Nothing, Optional ByRef blnClear As Boolean = False)
        FindResources()
        NewGame(prnt, mw, mh, m, gtype, tw, th, tClr, blnClear)
    End Sub

    Public Sub New(ByVal prnt As Control, ByRef dif As String, Optional ByRef gtype As Byte = 0, Optional ByRef tw As Short = 16, Optional ByRef th As Short = 16, Optional ByRef tClr As Color = Nothing, Optional ByRef blnClear As Boolean = False)
        FindResources()
        SetDifficulty(dif, prnt, gtype, tw, th, tClr, blnClear)
    End Sub

    Private Sub FindResources()
        Dim fileloc As New DirectoryInfo(Application.StartupPath)
        Do
            Dim folders() As DirectoryInfo = fileloc.GetDirectories()
            Dim filefolder As DirectoryInfo = folders.FirstOrDefault(Function(x) x.Name = "Content")
            If filefolder IsNot Nothing Then
                fnContent = filefolder.FullName & "\"
                fnAnimations = fnContent & "\Animations"
                Exit Do
            Else
                If fileloc.Parent Is Nothing Then
                    System.IO.Directory.CreateDirectory(Application.StartupPath)
                    Exit Do
                End If
                fileloc = fileloc.Parent
            End If
        Loop
    End Sub

    Private Sub NewGame(ByVal prnt As Control, ByRef mw As Short, ByRef mh As Short, ByRef m As Short, Optional ByRef gtype As Byte = 0, Optional ByRef tw As Short = 16, Optional ByRef th As Short = 16, Optional ByRef tClr As Color = Nothing, Optional ByRef blnClear As Boolean = False)
        Dim blnChange As Boolean = False
        Dim PGameType As Byte = gtype
        If mw <> W Or mh <> H Or MineCount <> m Then blnChange = True
        If tClr <> Nothing Then clr = tClr
        If clr = Nothing Then clr = Color.FromArgb(255, 153, 217, 234) Else clr = tClr

        If blnClear = False Then FillAnimationsArrays()
        blnGameOver = False
        blnVictory = False
        GameType = gtype
        W = mw
        H = mh
        TileW = tw
        TileH = th
        If gtype = 1 Then iExt = TileH / 2 Else iExt = 0
        MineCount = m
        If MineCount >= W * H Then MineCount = (W * H) - 1
        If blnClear = False Then
            GenerateAnimationsList()
            SetSmileyIndex(1)
            If LoadGraphics() = False Then Exit Sub
            Randomize()
        Else
            If Not img Is Nothing Then img.Dispose()
            If Not img Is Nothing Then img = Nothing
            If Not rst Is Nothing Then rst.Dispose()
            If Not rst Is Nothing Then rst = Nothing
            If Not animationTimer Is Nothing Then animationTimer.Dispose()
            If Not animationTimer Is Nothing Then animationTimer = Nothing
            pMF = Nothing
            Tiles = Nothing
            blnMouseDownMatch = False
            blnGameOver = False
            blnEnableQuestion = True
            blnFirstMove = False
        End If
        ReDim Tiles(W - 1, H - 1, 1)
        If Not img Is Nothing Then img.Dispose()
        img = New Bitmap(W * TileW, H * TileH + iExt)
        If Not TimeTaken Is Nothing Then TimeTaken.Dispose()
        TimeTaken = New Bitmap(Numbers(0).Width * 3, Numbers(0).Height)
        If Not MinesLeft Is Nothing Then MinesLeft.Dispose()
        MinesLeft = New Bitmap(Numbers(0).Width * 3, Numbers(0).Height)
        For r = 0 To W - 1
            For c = 0 To H - 1
                Tiles(r, c, 0) = ButtonUpIndex
            Next
        Next
        pMF = New PMouse

        If blnClear = False Then
            If TypeOf prnt Is Form Then
                Dim form1 As Form = prnt
                form1.MaximizeBox = False
                form1.MinimizeBox = False
            End If
            panGame = New Panel
            panGame.Parent = prnt
            panGame.BackColor = Color.Black
            panGame.BorderStyle = BorderStyle.Fixed3D
            panBar = New Panel
            panBar.Parent = prnt
            panBar.BorderStyle = BorderStyle.Fixed3D
            pReset = New PictureBox
            pMinefield = New PictureBox
            pMinesLeft = New PictureBox
            pTimeTaken = New PictureBox
            pReset.Parent = panBar
            pMinefield.Parent = panGame
            pMinesLeft.Parent = panBar
            pTimeTaken.Parent = panBar
            pMinesLeft.Left = 5
            pMinefield.Top = 0
            pMinefield.Left = 0
            pMinesLeft.Width = MinesLeft.Width
            pTimeTaken.Width = TimeTaken.Width
            pMinesLeft.Height = MinesLeft.Height
            pTimeTaken.Height = TimeTaken.Height
            SetUpMenu(prnt)
            SetToolTip(True)
            AddHandler pMinefield.MouseMove, AddressOf pMineField_MouseMove
            AddHandler pMinefield.MouseDown, AddressOf pMineField_MouseDown
            AddHandler pMinefield.MouseUp, AddressOf pMineField_MouseUp
            AddHandler pReset.MouseDown, AddressOf pReset_MouseDown
            AddHandler pReset.MouseUp, AddressOf pReset_MouseUp
            AddHandler objToolTip.Popup, AddressOf objToolTip_Popup
        End If

        AnimateSequence(0)

        panGame.Left = BorderBuffer
        panGame.Width = mw * TileW + 4
        panGame.Height = mh * TileH + 4 + iExt
        panBar.Left = BorderBuffer
        panBar.Top = BorderBuffer + 17
        panBar.Width = panGame.Width
        panBar.Height = ResetH + 10
        panGame.Top = panBar.Bottom + 5

        pMinefield.BorderStyle = BorderStyle.None
        pMinefield.Width = mw * tw
        pMinefield.Height = (mh * th) + (th / 2)
        panGame.Width = pMinefield.Width + 4
        panGame.Height = pMinefield.Height + iExt - 4

        pReset.Width = MineStates(ResetIndex).Width
        pReset.Height = MineStates(ResetIndex).Height
        pReset.Top = 5
        pReset.Left = (panBar.Width / 2) - (pReset.Width / 2)
        pMinesLeft.Top = 5
        pTimeTaken.Top = 5
        pTimeTaken.Left = panBar.Width - pTimeTaken.Width - 9

        Dim containerz As ContainerControl = pMinefield.GetContainerControl
        containerz.Width = panGame.Width + (BorderBuffer * 4)
        containerz.Height = panGame.Bottom + 42
        If blnChange = True Then
            containerz.Top = (Screen.PrimaryScreen.WorkingArea.Height / 2) - (containerz.Height / 2)
            containerz.Left = (Screen.PrimaryScreen.WorkingArea.Width / 2) - (containerz.Width / 2)
        End If
        MeLeft = containerz.Left
        MeTop = containerz.Top
        MeWidth = containerz.Width
        MeHeight = containerz.Height

        UpdateMinesLeft(0, MineCount)
        TimeTakenCounter = -1
        UpdateTimeTaken()
        DrawImg()
        If Not timetakenTimer Is Nothing Then
            timetakenTimer.Enabled = False
            timetakenTimer.Dispose()
        End If
        If Not prnt Is Nothing Then prnt = Nothing
    End Sub

    Private Sub FillAnimationsArrays()
        AList = New List(Of Animation)
        AList.Add(SmileyUpA)
        AList.Add(SmileyDownA)
        AList.Add(SmileyGameoverA)
        AList.Add(SmileyVictoryA)

        BList = New List(Of List(Of Bitmap))
        BList.Add(SmileyUp)
        BList.Add(SmileyDown)
        BList.Add(SmileyGameover)
        BList.Add(SmileyVictory)
    End Sub

    Private Sub GenerateAnimationsList()
        AnimationsList = New List(Of String)
        Dim di As New System.IO.DirectoryInfo(fnAnimations)
        Dim arrFi() As System.IO.FileInfo = di.GetFiles
        Dim fi As System.IO.FileInfo
        For Each fi In arrFi
            If fi.Extension <> ".jpg" And fi.Extension <> ".bmp" And fi.Extension <> ".png" Then Continue For
            AnimationsList.Add(fi.Name)
        Next
    End Sub
#End Region

#Region "Miscellaneous Form Functions"
    Private Sub SetToolTip(ByVal blnShow As Boolean, Optional ByRef intDelay As Integer = 1, Optional ByRef intLength As Integer = 10, Optional ByRef intReshowDelay As Integer = 1)
        ToolTipDisplay = blnShow
        objToolTip.InitialDelay = intDelay * 1000
        objToolTip.AutoPopDelay = intLength * 1000
        objToolTip.ReshowDelay = intReshowDelay * 1000
    End Sub

    Private Sub objToolTip_Popup(ByVal sender As Object, ByVal e As System.Windows.Forms.PopupEventArgs)
        e.Cancel = Not (ToolTipDisplay)
    End Sub
#End Region

#Region "Game Logic"
    Private Sub GameOver(ByRef x As Short, ByRef y As Short)
        For r = 0 To W - 1
            For c = 0 To H - 1
                If Tiles(r, c, 0) = FlagIndex And Tiles(r, c, 1) <> MineIndex Then
                    Using gx As Graphics = Graphics.FromImage(img)
                        gx.DrawImage(MineStates(MineIndex), r * TileW, c * TileH + CheckExt(x), TileW, TileH)
                        gx.DrawImage(MineStates(CrossIndex), r * TileW, c * TileH + CheckExt(x), TileW, TileH)
                    End Using
                ElseIf Tiles(r, c, 1) = MineIndex Then
                    SetTileIndex(r, c, MineIndex)
                End If

            Next
        Next
        timetakenTimer.Dispose()
        AnimateSequence(2)
        blnGameOver = True
        timetakenTimer.Dispose()
        FillColorAllSpaces(img, Color.Red, clr, x * TileW, y * TileH + CheckExt(x), TileW, TileH)
        pMinefield.Image = img
    End Sub

    Private Sub Victory()
        For r = 0 To W - 1
            For c = 0 To H - 1
                If Tiles(r, c, 0) = ButtonUpIndex Or Tiles(r, c, 0) = QuestionIndex Then SetTileIndex(r, c, FlagIndex)
            Next
        Next
        If CInt(MinesLeftCounter) <> 0 Then UpdateMinesLeft(0, 0)
        timetakenTimer.Enabled = False
        timetakenTimer.Dispose()
        blnGameOver = True
        blnVictory = True
        AnimateSequence(3)
        Exit Sub
        'If MBox.NewMsgBox("You're Winner! Liek, OMG WTF BBQ yay! shane is terrible at serious sam. By the way, do you have any feces to eat?", "yay!", MBox.MFont("Comic Sans MS", 10, FontStyle.Regular, GraphicsUnit.Pixel, Brushes.Black), New SolidBrush(clr), New BMB.MsgBoxBtns(MBox.btn("OK", "Yay!"), MBox.btn("No", "uh no!"), MBox.btn("troll", "um, shane is bad at serious sam")), MBox.NewIcon(SmileyAnimations(SIndex).Item(0), 60, 60)) = "OK" Then
        'MsgBox("omg... you did it!")
        'End If
    End Sub

    Public Sub CalculateResult(ByRef x As Short, ByRef y As Short)
        If blnFirstMove = False Then
            GenerateMinefield(x, y)
            blnFirstMove = True
            timetakenTimer = New Timer
            timetakenTimer.Interval = 1000
            AddHandler timetakenTimer.Tick, AddressOf timetaken_Tick
            timetakenTimer.Enabled = True
        End If
        If Tiles(x, y, 1) = MineIndex Then
            GameOver(x, y)
        Else
            CheckSpaces(x, y, New List(Of Short))
            pMinefield.Image = img
        End If
        For r = 0 To W - 1
            For c = 0 To H - 1
                If Tiles(r, c, 1) <> MineIndex Then
                    If Tiles(r, c, 0) > 9 Then
                        Exit Sub
                    End If
                End If
            Next
        Next
        Victory()
    End Sub

    Private Sub CheckSpaces(ByRef x As Short, ByRef y As Short, ByRef CheckedSpaces As List(Of Short))
        Dim ext As Short = x Mod 2
        Dim iCount As Short = 0
        For r = x - 1 To x + 1
            For c = y - 1 To y + 1
                If (r = x And c = y) Or r < 0 Or c < 0 Or r = W Or c = H Then Continue For
                If GameType = 1 Then If ext = 1 Then If (c = y - 1 And r <> x) Then Continue For Else  Else If (c = y + 1 And r <> x) Then Continue For
                If Tiles(r, c, 1) = MineIndex Then iCount += 1
            Next
        Next
        If iCount > 0 Then
            SetTileIndex(x, y, iCount)
        Else
            SetTileIndex(x, y, 0)
            For r = x - 1 To x + 1
                For c = y - 1 To y + 1
                    If (r = x And c = y) Or r < 0 Or c < 0 Or r = W Or c = H Then Continue For
                    If GameType = 1 Then If ext = 1 Then If (c = y - 1 And r <> x) Then Continue For Else  Else If (c = y + 1 And r <> x) Then Continue For
                    Dim index As Short = (W * c) + r
                    If CheckedSpaces.Contains(index) = False Then
                        CheckedSpaces.Add(index)
                        CheckSpaces(r, c, CheckedSpaces)
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub GenerateMinefield(ByRef x As Short, ByRef y As Short)
        Dim iSides As Short = 0
        Dim ext As Short = CheckExt(x)
        Dim lstNum As New List(Of Integer)
        Dim lstSafe As New List(Of Short)

        For r = x - 1 To x + 1
            For c = y - 1 To y + 1
                If (r = x And c = y) Or r < 0 Or c < 0 Or r = W Or c = H Then Continue For
                If GameType = 1 Then If ext = 1 Then If (c = y - 1 And r <> x) Then Continue For Else  Else If (c = y + 1 And r <> x) Then Continue For
                lstSafe.Add((W * c) + r)
            Next
        Next

        Dim iCount As Short = (W * H)
        lstNum = Enumerable.Range(0, iCount).ToList
        lstNum.Remove((W * y) + x)

        Dim iMin As Short = Math.Min(lstSafe.Count, lstNum.Count - MineCount)
        If lstSafe.Count > 0 And iMin > 0 Then
            Dim num As Short = lstSafe(RandomInteger(0, lstSafe.Count))
            lstNum.Remove(num)
            iMin -= 1
        End If

        For i = 0 To iMin - 1
            If i >= lstSafe.Count Then Exit For
            Dim num As Short = RandomInteger(0, 2)
            If num = 0 Then lstNum.Remove(lstSafe(i))
        Next

        For i = 0 To MineCount - 1
            Dim num As Short = lstNum(RandomInteger(0, lstNum.Count))
            Dim c As Short = Math.Floor(num / W)
            Dim r As Short = num Mod W
            Tiles(r, c, 1) = MineIndex
            lstNum.Remove(num)
        Next
    End Sub

    Private Sub UpdateMinesLeft(ByRef num As Short, Optional ByRef initialnum As Short = -1)
        If initialnum <> -1 Then MinesLeftCounter = initialnum Else MinesLeftCounter = CStr(CInt(MinesLeftCounter) + num)
        If CInt(MinesLeftCounter > -1) Then
            If MinesLeftCounter.Length < 2 Then MinesLeftCounter = MinesLeftCounter.Insert(0, "0")
            If MinesLeftCounter.Length < 3 Then MinesLeftCounter = MinesLeftCounter.Insert(0, "0")
        Else
            If MinesLeftCounter.Length = 2 Then MinesLeftCounter = MinesLeftCounter.Insert(1, "0")
        End If
        Using gx As Graphics = Graphics.FromImage(MinesLeft)
            For i = 0 To 2
                If MinesLeftCounter.Chars(i) = "-" Then gx.DrawImage(Numbers(10), 0, 0) Else gx.DrawImage(Numbers(CInt(MinesLeftCounter.Substring(i, 1))), Numbers(0).Width * i, 0)
            Next
        End Using
        pMinesLeft.Image = MinesLeft
    End Sub

    Private Sub UpdateTimeTaken()
        If TimeTakenCounter > 999 Then
            timetakenTimer.Enabled = False
            timetakenTimer.Dispose()
            Exit Sub
        End If
        TimeTakenCounter += 1
        Dim sTimeTaken As String = CStr(TimeTakenCounter)
        If sTimeTaken.Length < 2 Then sTimeTaken = sTimeTaken.Insert(0, "0")
        If sTimeTaken.Length < 3 Then sTimeTaken = sTimeTaken.Insert(0, "0")
        Using gx As Graphics = Graphics.FromImage(TimeTaken)
            For i = 0 To 2
                gx.DrawImage(Numbers(CInt(sTimeTaken.Substring(i, 1))), Numbers(0).Width * i, 0)
            Next
        End Using
        pTimeTaken.Image = TimeTaken
    End Sub

    Private Sub timetaken_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        UpdateTimeTaken()
    End Sub
#End Region

#Region "General Functions"

    Private Function RandomInteger(ByRef HighValue As Short, ByRef LowValue As Short) As Short
        Return Int((HighValue - LowValue) * Rnd() + LowValue)
    End Function

    Private Function CheckExt(ByRef x As Short)
        If GameType = 1 Then If x Mod 2 = 1 Then Return iExt Else Return 0 Else Return 0
    End Function
#End Region

#Region "Drawing"
    Private Sub DrawImg()
        Using gx As Graphics = Graphics.FromImage(img)
            For c = 0 To H - 1
                For r = 0 To W - 1
                    gx.DrawImage(MineStates(Tiles(r, c, 0)), r * TileW, c * TileH + CheckExt(r), TileW, TileH)
                Next
            Next
        End Using
        pMinefield.Image = img
    End Sub

    Private Sub DrawToFrame(ByRef x As Short, ByRef y As Short)
        Using gx As Graphics = Graphics.FromImage(img)
            gx.DrawImage(MineStates(Tiles(x, y, 0)), x * TileW, y * TileH + CheckExt(x), TileW, TileH)
        End Using
    End Sub

    Private Sub SetTileIndex(ByRef x As Short, ByRef y As Byte, ByRef i As Short)
        Tiles(x, y, 0) = i
        Using gx As Graphics = Graphics.FromImage(img)
            gx.DrawImage(MineStates(Tiles(x, y, 0)), x * TileW, y * TileH + CheckExt(x), TileW, TileH)
        End Using
    End Sub
#End Region

#Region "Menu Events"
    Private PrevioussMenuItem As Short = -1
    Private blnCustomMap As Boolean = False
    Private arrIndexes As List(Of Short)
    Private txtSearchText As String = ""
    Private lstBox As ListBox
    Private AnimationsList As List(Of String)
    Private panAnimations As List(Of Panel)
    Private pPreview As List(Of PictureBox)
    Private tmrPreview As List(Of Timer)
    Private frmAnimations_ShowPreviewBackground As Boolean
    Private frmAnimations_ShowFrameBackground As Boolean
    Private frmAnimations_ChangesMade As Boolean
    Private frmAnimations_PrevIndex As Short = -1
    Private frmAnimations_PreviewRepeatInterval As List(Of Integer)
    Private frmAnimations_PreviewRepeatDefault As Integer = 2000
    Private AListClone() As Animation

    Private Sub SetUpMenu(ByVal prnt As Control)
        MainMenu = New MenuStrip
        MainMenu.Parent = prnt

        Dim TempItem As New ToolStripMenuItem("Game")
        TempItem.DropDownItems.Add("Board Type")
        TempItem.DropDownItems.Add("Beginner", Nothing, AddressOf mnuGame_Click)
        TempItem.DropDownItems.Add("Intermediate", Nothing, AddressOf mnuGame_Click)
        TempItem.DropDownItems.Add("Expert", Nothing, AddressOf mnuGame_Click)
        TempItem.DropDownItems.Add("Custom", Nothing, AddressOf mnuGame_Click)
        TempItem.DropDownItems(1).ToolTipText = "9 x 9, 10 mines"
        TempItem.DropDownItems(2).ToolTipText = "16 x 16, 40 mines"
        TempItem.DropDownItems(3).ToolTipText = "30 x 16, 100 mines"
        TempItem.DropDownItems(4).ToolTipText = "Custon map size and mine count"
        For i = 0 To TempItem.DropDownItems.Count - 1
            Dim mnuItem As ToolStripMenuItem = TempItem.DropDownItems.Item(i)
            If i = 0 Then
                mnuItem.DropDownItems.Add("0 - Standard", Nothing, AddressOf mnuGame_Click)
                Dim mnuItemTemp As ToolStripMenuItem = mnuItem.DropDownItems(0)
                mnuItemTemp.Checked = True
                mnuItem.DropDownItems.Add("1 - Hexagonal", Nothing, AddressOf mnuGame_Click)
            End If
            If i = 1 Then mnuItem.Checked = True
        Next
        MainMenu.Items.Add(TempItem)

        TempItem = New ToolStripMenuItem("Options")
        Dim mnuTemp2 As New ToolStripMenuItem("Animations")
        For i = 0 To AnimationsList.Count - 1
            mnuTemp2.DropDownItems.Add(AnimationsList(i).Substring(0, AnimationsList(i).LastIndexOf(".")), Nothing, AddressOf mnuAnimations_Click)
            mnuTemp2.DropDownItems(i).Tag = i
            If i = SIndex Then
                Dim mnuTemp3 As ToolStripMenuItem = mnuTemp2.DropDownItems(i)
                mnuTemp3.Checked = True
            End If
        Next
        TempItem.DropDownItems.Add(mnuTemp2)
        Dim mnuEdit As New ToolStripMenuItem("Edit")
        mnuEdit.DropDownItems.Add("Animations", Nothing, AddressOf mnuEdit_Click)
        mnuEdit.DropDownItems.Add("Tile Color", Nothing, AddressOf mnuEdit_Click)
        mnuEdit.DropDownItems(0).ToolTipText = "Edit victory/gameover animations"
        TempItem.DropDownItems.Add(mnuEdit)
        MainMenu.Items.Add(TempItem)
    End Sub

    Private Sub mnuAnimations_Click(ByVal mnuItem As ToolStripMenuItem, ByVal e As System.EventArgs)
        If mnuItem.Checked = False Then
            For i = 0 To mnuItem.GetCurrentParent.Items.Count - 1
                Dim mnuTemp As ToolStripMenuItem = mnuItem.GetCurrentParent.Items(i)
                If mnuTemp.Text = mnuItem.Text Then Continue For
                If mnuTemp.Checked = True Then mnuTemp.Checked = False
            Next
        End If
        If mnuItem.Checked = False Then
            mnuItem.Checked = True
            Dim lolz As Short = mnuItem.Tag
            Dim currentIndex As Byte = AIndex
            LoadAnimation(fnAnimations & "\" & AnimationsList(mnuItem.Tag))
            AnimateSequence(currentIndex)
        End If
    End Sub

    Private Sub mnuEdit_Click(ByVal mnuItem As ToolStripMenuItem, ByVal e As System.EventArgs)
        Select Case mnuItem.Text
            Case "Animations"
                EditAnimations()
            Case "Tile Color"
                Dim yolo As New ColorDialog
                yolo.ShowDialog()
                If yolo.Color = clr Then Exit Sub
                clr = yolo.Color

                Dim sourceRect As New Rectangle(0, 0, DefaultTileW, DefaultTileH)
                Dim clrDarker As Color = DarkerColor(clr)
                Dim clrBorder As Color = BorderColor(clr)
                SetTileColors(False, clr, clrDarker, clrBorder)
                For i = 1 To 9
                    Dim newImg As Bitmap
                    If i = 9 Then
                        newImg = Image.FromFile(fnContent & "\Mine.png")
                    Else
                        newImg = Image.FromFile(fnContent & "\" & i & ".png")
                    End If
                    MineStates(i) = DrawToBaseImage(MineStates(0).Clone(sourceRect, Imaging.PixelFormat.DontCare), newImg, i, DefaultTileW, DefaultTileH)
                Next
                SetTileColors(True, clr, clrDarker, clrBorder)
                For i = ButtonUpIndex + 1 To ButtonUpIndex + 2
                    Dim newImg As Bitmap
                    If i = ButtonUpIndex + 1 Then
                        newImg = Image.FromFile(fnContent & "\Flag.png")
                        FlagIndex = i
                    Else
                        newImg = Image.FromFile(fnContent & "\QuesitonMark.png")
                        QuestionIndex = i
                    End If
                    MineStates(i) = DrawToBaseImage(MineStates(ButtonUpIndex).Clone(sourceRect, Imaging.PixelFormat.DontCare), newImg, i)
                Next
                DrawImg()
        End Select

    End Sub

    Private Sub EditAnimations()
        Dim frmAnimations As New Form()
        Dim sideBuffer As Short = 10
        CreateAListClone()

        frmAnimations.Text = "Edit Animations"
        frmAnimations.MaximizeBox = False
        frmAnimations.MinimizeBox = False
        frmAnimations_ShowFrameBackground = False
        frmAnimations_ShowPreviewBackground = True
        frmAnimations_ChangesMade = False

        Dim newTextBox As New TextBox()
        newTextBox.Parent = frmAnimations
        newTextBox.Name = "txtSearchAnimations"
        newTextBox.Left = sideBuffer
        newTextBox.Top = sideBuffer
        newTextBox.Height = 20
        newTextBox.Width = 135
        AddHandler newTextBox.KeyUp, AddressOf frmAnimations_Textbox_KeyUp
        objToolTip.SetToolTip(newTextBox, "Type a phrase to instant-search the below list of Animation files.")

        lstBox = New ListBox()
        lstBox.Parent = frmAnimations
        lstBox.Left = sideBuffer
        lstBox.Top = newTextBox.Bottom + sideBuffer
        lstBox.Height = 66
        lstBox.Width = 135
        PopulateListbox(lstBox)
        lstBox.SelectedIndex = SIndex
        AddHandler lstBox.SelectedIndexChanged, AddressOf frmAnimations_ListBox_SelectedIndexChanged
        objToolTip.SetToolTip(lstBox, "List of Animation files in the Content\Animations directory.")

        Dim panRadioPreviewControl As New Panel()
        panRadioPreviewControl.Parent = frmAnimations
        panRadioPreviewControl.Name = "panRadioPreviewControl"
        panRadioPreviewControl.Left = lstBox.Right + sideBuffer
        panRadioPreviewControl.Top = newTextBox.Top
        panRadioPreviewControl.Height = lstBox.Bottom - newTextBox.Top - 10
        panRadioPreviewControl.Width = 105
        For i = 0 To 2
            Dim newRB As New RadioButton()
            newRB.Parent = panRadioPreviewControl
            newRB.AutoSize = True
            newRB.Top = 5 + (i * 20)
            newRB.Left = sideBuffer
            newRB.Name = "radCheckbox" & i
            If i = 0 Then newRB.Text = "Repeat All" Else If i = 1 Then newRB.Text = "Custom" Else If i = 2 Then newRB.Text = "Repeat None"
            If i = 0 Then newRB.Checked = True
            AddHandler newRB.CheckedChanged, AddressOf EditAnimations_RepeatRadioCheckChanged
        Next

        panRadioPreviewControl.BorderStyle = BorderStyle.FixedSingle
        Dim panCheckboxResetBackgroundDisplay As New Panel
        panCheckboxResetBackgroundDisplay.Parent = frmAnimations
        panCheckboxResetBackgroundDisplay.Name = "panCheckboxResetBackgroundDisplay"
        panCheckboxResetBackgroundDisplay.Left = panRadioPreviewControl.Right + sideBuffer
        panCheckboxResetBackgroundDisplay.Top = newTextBox.Top
        panCheckboxResetBackgroundDisplay.Width = 135
        panCheckboxResetBackgroundDisplay.Height = panRadioPreviewControl.Height
        Dim newLabel As New Label()
        newLabel.Parent = panCheckboxResetBackgroundDisplay
        newLabel.AutoSize = True
        newLabel.Text = "Display Button BG Image"
        newLabel.Left = sideBuffer - 4
        newLabel.Top = 5
        For i = 0 To 1
            Dim newCB As New CheckBox()
            newCB.Parent = panCheckboxResetBackgroundDisplay
            newCB.Tag = i
            newCB.AutoSize = True
            newCB.Left = sideBuffer
            newCB.Top = newLabel.Bottom + 5 + (i * 20)
            If i = 0 Then newCB.Text = "Preview Animation" Else If i = 1 Then newCB.Text = "Frame Image"
            If i = 0 Then If frmAnimations_ShowPreviewBackground = True Then newCB.Checked = True
            If i = 1 Then If frmAnimations_ShowFrameBackground = True Then newCB.Checked = True
            AddHandler newCB.CheckedChanged, AddressOf EditAnimations_ShowBackgroundCheckedChanged
        Next

        panCheckboxResetBackgroundDisplay.BorderStyle = BorderStyle.FixedSingle

        newTextBox = New TextBox()
        newTextBox.Parent = frmAnimations
        newTextBox.Name = "txtRenameAnimation"
        newTextBox.Left = lstBox.Left
        newTextBox.Top = lstBox.Bottom - 10 + sideBuffer
        newTextBox.Height = 20
        newTextBox.Width = 150
        newTextBox.BackColor = frmAnimations.BackColor
        newTextBox.Text = AnimationsList(SIndex)
        objToolTip.SetToolTip(newTextBox, "Change animation file name")
        AddHandler newTextBox.KeyDown, AddressOf mnuTextbox_AnimationNameChanged

        Dim newButton As New Button()
        newButton.Parent = frmAnimations
        newButton.AutoSize = True
        newButton.Left = newTextBox.Right + sideBuffer
        newButton.Top = newTextBox.Top
        newButton.Text = "Save"
        objToolTip.SetToolTip(newButton, "Save changes to animation sequence")
        AddHandler newButton.Click, AddressOf frmAnimations_SaveClick

        panAnimations = New List(Of Panel)
        For i = 0 To 3
            Dim newPanel As New Panel
            newPanel = New Panel
            newPanel.Parent = frmAnimations
            newPanel.Tag = i
            newPanel.Width = 395
            newPanel.Height = 80
            newPanel.Left = sideBuffer
            newPanel.Top = newTextBox.Bottom + 15 + (newPanel.Height * i) + (15 * i)
            newPanel.BorderStyle = BorderStyle.FixedSingle
            newPanel.AutoScroll = True
            newPanel.SetAutoScrollMargin(5, 5)
            panAnimations.Add(newPanel)
            newLabel = New Label()
            newLabel.Parent = frmAnimations
            If i = 0 Then newLabel.Text = "Normal" Else If i = 1 Then newLabel.Text = "Button Pressed" Else If i = 2 Then newLabel.Text = "Gameover" Else If i = 3 Then newLabel.Text = "Victory"
            newLabel.AutoSize = True
            newLabel.Top = newPanel.Top - newLabel.Height
            newLabel.Left = newPanel.Left - 3
        Next
        LoadPanels()

        frmAnimations.Width = panAnimations(0).Right + sideBuffer + 16
        frmAnimations.Height = panAnimations(panAnimations.Count - 1).Bottom + sideBuffer + 38
        AddHandler frmAnimations.Load, AddressOf frmAnimations_Load
        AddHandler frmAnimations.FormClosing, AddressOf frmAnimations_Closing
        frmAnimations.ShowDialog()
    End Sub

    Private Sub frmAnimations_Load(ByVal frmSender As Form, ByVal e As System.EventArgs)
        frmSender.Left = MeLeft + (MeWidth / 2) - (frmSender.Width / 2)
        frmSender.Top = MeTop + (MeHeight / 2) - (frmSender.Height / 2)
    End Sub

    Private Sub frmAnimations_Closing(ByVal frmSender As Form, ByVal e As System.EventArgs)
        For i = 0 To tmrPreview.Count - 1
            pPreview(i) = Nothing
            tmrPreview(i).Enabled = False
            tmrPreview(i).Dispose()
            panAnimations(i).Dispose()
        Next
        pPreview = Nothing
        tmrPreview = Nothing
        panAnimations = Nothing
        AListClone = Nothing
        frmSender.Dispose()
    End Sub

    Private Sub mnuTextbox_AnimationNameChanged(ByVal txtSender As TextBox, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Oem5 Or e.KeyCode = Keys.Oemtilde Or e.KeyCode = Keys.OemQuestion Or e.KeyCode = Keys.Oemplus Or (e.Shift = True AndAlso (e.KeyCode = Keys.D8 Or e.KeyCode = Keys.OemPeriod Or e.KeyCode = Keys.Oemcomma Or e.KeyCode = Keys.OemQuotes Or e.KeyCode = Keys.OemSemicolon)) Then
            e.SuppressKeyPress = True
            Exit Sub
        End If
        frmAnimations_ChangesMade = True
    End Sub

    Private Sub frmAnimations_SaveClick(ByVal btnSender As Button, ByVal e As MouseEventArgs)
        If frmAnimations_ChangesMade = False Then Exit Sub
        If MessageBox.Show("Save Changes?", "uh do it?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
        AList.Clear()
        AList = New List(Of Animation)
        For i = 0 To AListClone.Count - 1
            Dim newAnimation As New Animation
            newAnimation.FrameIndex = 0
            newAnimation.InfiniteRepeat = AListClone(i).InfiniteRepeat
            newAnimation.RepeatCounter = AListClone(i).RepeatCounter
            newAnimation.FrameTimer = New List(Of Short)
            For i2 = 0 To AListClone(i).FrameTimer.Count - 1
                newAnimation.FrameTimer.Add(AListClone(i).FrameTimer(i2))
            Next
            AList.Add(newAnimation)
        Next
        objAnimation = AList(AIndex)

        Dim sFile As String = fnAnimations & "\" & "Sequences.txt"
        If SF.CheckFileExists(sFile) = False Then Exit Sub

        Dim arrLines As List(Of String) = SF.GetArrayLinesFromStream(sFile, True, True)
        If arrLines Is Nothing Then Exit Sub
        Dim blnFound As Boolean = False
        For i = 0 To arrLines.Count - 1
            Dim sItem As String = arrLines(i).Substring(0, arrLines(i).IndexOf(":"))
            If AnimationsList(SIndex) = sItem Then
                arrLines.RemoveAt(i)
                blnFound = True
                Exit For
            End If
        Next

        Dim NewSmileyW As Short = SmileyW
        Dim NewSmileyH As Short = SmileyH
        If blnFound = False Then
            Do
                Dim sResponse As String = InputBox("Enter the size of individual frames" & vbNewLine & "Default: " & vbNewLine & "   17x17 pixels" & vbNewLine & "   Minimum 8x8" & vbNewLine & "   Maximum 60x60" & vbNewLine & " in format: #x#", "uh do it!", "17x17").Trim
                Dim sUResponse As String = sResponse.ToUpper
                If sUResponse.Contains("X") = False Then If MessageBox.Show("Error: Must contain " & Chr(34) & "x" & Chr(34) & "between two numbers." & vbNewLine & vbNewLine & "Try again? (Will default to 17x17 otherwise)", "stop messing up!", MessageBoxButtons.YesNo, MessageBoxIcon.Error) = DialogResult.No Then Exit Do Else Continue Do
                If sUResponse.StartsWith("X") Or sUResponse.EndsWith("X") Then If MessageBox.Show("Error: Must contain 2 numbers, one before and one after the x to specify width and height of individual frames." & vbNewLine & vbNewLine & "Try again? (Will default to 17x17 otherwise)", "stop messing up!", MessageBoxButtons.YesNo, MessageBoxIcon.Error) = DialogResult.No Then Exit Do Else Continue Do
                Dim arrNum As New List(Of String)
                arrNum.Add(sResponse.Substring(0, sUResponse.IndexOf("X")))
                arrNum.Add(sResponse.Substring(sUResponse.IndexOf("X") + 1, sResponse.Length - sUResponse.IndexOf("X") - 1))
                For i = 0 To 1
                    For i2 = 0 To arrNum(i).Length - 1
                        If arrNum(i).ToUpper.Chars(i2) <> "X" AndAlso (Asc(arrNum(i).Chars(i2)) < 48 Or Asc(arrNum(i).Chars(i2)) > 57) Then If MessageBox.Show("Error: Must contain only numbers before and after the x. Invalid character typed: " & Chr(34) & arrNum(i).Chars(i2) & Chr(34) & vbNewLine & vbNewLine & "Try again? (Will default to 17x17 otherwise)", "stop messing up!", MessageBoxButtons.YesNo, MessageBoxIcon.Error) = DialogResult.No Then Exit Do Else Continue Do
                    Next
                    If Val(arrNum(i)) < 8 Or Val(arrNum(i)) > 60 Then If MessageBox.Show("Error: Width/Height must each be between 8 and 60." & vbNewLine & vbNewLine & "Try again? (Will default to 17x17 otherwise)", "stop messing up!", MessageBoxButtons.YesNo, MessageBoxIcon.Error) = DialogResult.No Then Exit Do Else Continue Do
                Next
                NewSmileyW = CShort(arrNum(0))
                NewSmileyH = CShort(arrNum(1))
                Exit Do
            Loop
        End If

        Dim sName As String = ""
        If btnSender.FindForm.Controls.ContainsKey("txtRenameAnimation") Then
            Dim txtTemp As TextBox = btnSender.FindForm.Controls("txtRenameAnimation")
            sName = txtTemp.Text
            If AnimationsList(SIndex) <> txtTemp.Text Then If SF.RenameFile(fnAnimations & "\" & AnimationsList(SIndex), txtTemp.Text, True, True) = False Then Exit Sub
        Else
            sName = AnimationsList(SIndex)
        End If

        Dim strBuilder As New System.Text.StringBuilder
        strBuilder.Append(sName & ":" & NewSmileyW & "," & NewSmileyH & "|")
        For i = 0 To AListClone.Count - 1
            strBuilder.Append("[")
            strBuilder.Append(CStr(AListClone(i).InfiniteRepeat))
            strBuilder.Append(",")
            strBuilder.Append(AListClone(i).RepeatCounter)
            strBuilder.Append("]")
            For i2 = 0 To AListClone(i).FrameTimer.Count - 1
                strBuilder.Append(AListClone(i).FrameTimer(i2))
                strBuilder.Append(",")
            Next
        Next
        arrLines.Add(strBuilder.ToString)
        arrLines.Sort()
        frmAnimations_ChangesMade = False

        Dim sw As System.IO.StreamWriter = Nothing
        sw = SF.SW_CreateText(sFile, sw, True, True)
        If sw Is Nothing Then Exit Sub
        For i = 0 To arrLines.Count - 1
            sw.WriteLine(arrLines(i))
        Next
        sw.Close()
    End Sub

    Private Sub EditAnimations_ShowBackgroundCheckedChanged(ByVal chkSender As CheckBox, ByVal e As System.EventArgs)
        If chkSender.Tag = 0 Then
            If chkSender.Checked = True Then frmAnimations_ShowPreviewBackground = True
            If chkSender.Checked = False Then frmAnimations_ShowPreviewBackground = False
            For i = 0 To pPreview.Count - 1
                Dim index As Short = AListClone(i).FrameIndex
                If panAnimations(i).Controls.ContainsKey("chk" & i) Then
                    Dim tempCB As CheckBox = panAnimations(i).Controls("chk" & i)
                    If tempCB.Checked = True Then index -= 1
                End If
                If index < 0 Then index = AListClone(i).FrameTimer.Count - 1
                pPreview(i).Image = GetPreviewImg(BList(i).Item(index), 0)
            Next
        ElseIf chkSender.Tag = 1 Then
            If chkSender.Checked = True Then frmAnimations_ShowFrameBackground = True
            If chkSender.Checked = False Then frmAnimations_ShowFrameBackground = False
            For i = 0 To panAnimations.Count - 1
                Dim iCounter As Short = -1
                For Each controlz As Control In panAnimations(i).Controls
                    If TypeOf controlz Is PictureBox Then
                        Dim tempPB As PictureBox = controlz
                        If iCounter = -1 Then
                            iCounter = 0
                            Continue For
                        End If
                        tempPB.Image = GetPreviewImg(BList(i).Item(iCounter), 1)
                        iCounter += 1
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub EditAnimations_RepeatRadioCheckChanged(ByVal rdbSender As RadioButton, ByVal e As System.EventArgs)
        If rdbSender.Checked = False Or rdbSender.Text = "Custom" Then Exit Sub
        For i = 0 To panAnimations.Count - 1
            If panAnimations(i).Controls.ContainsKey("chk" & i) Then
                Dim tempCB As CheckBox = panAnimations(i).Controls("chk" & i)
                If rdbSender.Text = "Repeat All" Then tempCB.Checked = True Else tempCB.Checked = False
            End If
        Next
    End Sub

    Private Sub LoadPanels()
        frmAnimations_PreviewRepeatInterval = New List(Of Integer)
        If panAnimations(0).Controls.Count > 0 Then 'If controls already exist, don't delete but change them instead
            For i = 0 To 3
                Dim aniTemp As Animation = AListClone(i)
                Dim bmpTemp As List(Of Bitmap) = BList(i)
                Dim PBC As Short = 0
                Dim TBC As Short = 0
                Dim CC As Short = panAnimations(i).Controls.Count - 1
                Dim DC As Short = 0
                Dim IR As Short = 0
                For i2 = 0 To CC
                    Dim controlz As Control = panAnimations(i).Controls(i2 - DC)
                    If TypeOf controlz Is PictureBox Then
                        Dim PB As PictureBox = controlz
                        If PBC = 0 Then
                            PB.Image = GetPreviewImg(bmpTemp(0), 0)
                            IR = PB.Right
                            PBC += 1
                            Continue For
                        End If
                        If PBC = bmpTemp.Count + 1 Then
                            panAnimations(i).Controls.Remove(controlz)
                            DC += 1
                        Else
                            PB.Image = bmpTemp(PBC - 1)
                            PBC += 1
                        End If
                    ElseIf TypeOf controlz Is TextBox Then
                        Dim TB As TextBox = controlz
                        If TBC = 0 Then
                            frmAnimations_PreviewRepeatInterval.Add(frmAnimations_PreviewRepeatDefault)
                            TBC += 1
                            Continue For
                        End If
                        If TBC = aniTemp.FrameTimer.Count + 1 Then
                            panAnimations(i).Controls.Remove(controlz)
                            DC += 1
                        Else
                            TB.Text = aniTemp.FrameTimer(TBC - 1)
                            TBC += 1
                        End If
                    End If
                Next
                Dim iStart As Short = Math.Max(TBC, PBC - 1)
                For i2 = TBC - 1 To aniTemp.FrameTimer.Count - 1 'Create the missing PBs and TBs
                    EditAnimations_CreatePBandTB(i, i2, IR + 5, bmpTemp, aniTemp)
                Next
            Next
        Else 'No primary controls yet, create them within the panels
            pPreview = New List(Of PictureBox)
            tmrPreview = New List(Of Timer)
            For i = 0 To panAnimations.Count - 1
                Dim aniTemp As Animation = AListClone(i)
                Dim bmpTemp As List(Of Bitmap) = BList(i)
                Dim newPP As New PictureBox()
                newPP.Parent = panAnimations(i)
                newPP.Left = 5
                newPP.Top = 5
                newPP.Width = ResetW
                newPP.Height = ResetH
                newPP.SizeMode = PictureBoxSizeMode.CenterImage
                newPP.Image = GetPreviewImg(bmpTemp(0), 0)
                newPP.Tag = i
                pPreview.Add(newPP)

                Dim newTmr As New Timer()
                newTmr.Tag = i
                AddHandler newTmr.Tick, AddressOf tmrPreview_Tick
                tmrPreview.Add(newTmr)

                Dim newTB As New TextBox()
                newTB.Parent = panAnimations(i)
                newTB.Left = newPP.Left
                newTB.Top = newPP.Bottom + 5
                newTB.Height = 20
                newTB.Width = 30
                newTB.Text = frmAnimations_PreviewRepeatDefault
                newTB.Name = "tPreview" & i
                newTB.Tag = i
                frmAnimations_PreviewRepeatInterval.Add(frmAnimations_PreviewRepeatDefault)
                AddHandler newTB.KeyDown, AddressOf mnuTextbox_NumbersOnly
                AddHandler newTB.KeyUp, AddressOf mnuTextbox_FrameTimeChanged

                'Dim newCB As New CheckBox()
                'newCB.Parent = panAnimations(i)
                'newCB.Tag = i
                'newCB.Name = "chk" & i
                'newCB.Text = "R"
                'newCB.AutoSize = True
                'newCB.Checked = True
                'newCB.Left = newPP.Left
                'newCB.Top = newTB.Bottom + 5
                'AddHandler newCB.CheckedChanged, AddressOf EditAnimations_CheckChanged

                For i2 = 0 To bmpTemp.Count - 1
                    EditAnimations_CreatePBandTB(i, i2, newPP.Right + 5, bmpTemp, aniTemp)
                Next
            Next
        End If
        For i = 0 To AListClone.Count - 1 'Start preview animations
            Dim aniTemp As Animation = AListClone(i)
            tmrPreview(i).Interval = aniTemp.FrameTimer(0)
            tmrPreview(i).Enabled = True
        Next
    End Sub

    Private Sub EditAnimations_CreatePBandTB(ByRef i As Integer, ByRef i2 As Integer, ByRef IR As Short, ByRef bmpTemp As List(Of Bitmap), ByRef aniTemp As Animation)
        Dim newPB As New PictureBox()
        newPB.Parent = panAnimations(i)
        newPB.Top = 5
        newPB.Left = IR + 20 + (SmileyW * i2) + (20 * i2)
        newPB.Width = ResetW
        newPB.Height = ResetH
        newPB.SizeMode = PictureBoxSizeMode.CenterImage
        newPB.Image = GetPreviewImg(bmpTemp(i2), 1)
        newPB.Name = "p" & i & "s" & i2

        Dim newTB As New TextBox()
        newTB.Parent = panAnimations(i)
        newTB.Left = newPB.Left
        newTB.Top = newPB.Bottom + 5
        newTB.Height = 20
        newTB.Width = 30
        newTB.Text = aniTemp.FrameTimer(i2)
        newTB.Name = "t" & i & "s" & i2
        newTB.Tag = i2
        AddHandler newTB.KeyDown, AddressOf mnuTextbox_NumbersOnly
        AddHandler newTB.KeyUp, AddressOf mnuTextbox_FrameTimeChanged
    End Sub

    Private Function GetPreviewImg(ByVal bmpImg As Bitmap, ByVal iType As Byte)
        If iType = 0 Then
            If frmAnimations_ShowPreviewBackground = True Then
                Return DrawToBaseImage(MineStates(ResetIndex).Clone(New Rectangle(0, 0, ResetW, ResetH), Imaging.PixelFormat.DontCare), bmpImg, -1, DefaultResetW, DefaultResetH)
            Else
                Return bmpImg
            End If
        ElseIf iType = 1 Then
            If frmAnimations_ShowFrameBackground = True Then
                Return DrawToBaseImage(MineStates(ResetIndex).Clone(New Rectangle(0, 0, ResetW, ResetH), Imaging.PixelFormat.DontCare), bmpImg, -1, DefaultResetW, DefaultResetH)
            Else
                If ResetW <> DefaultResetW Or ResetH <> DefaultResetH Then
                    Dim newBmp As New Bitmap(ResetW, ResetH) 'bmpImg.Clone(New Rectangle(0, 0, SmileyW, SmileyH), Imaging.PixelFormat.DontCare)
                    Using gx As Graphics = Graphics.FromImage(newBmp)
                        gx.FillRectangle(New SolidBrush(Color.FromArgb(255, 255, 0, 128)), New Rectangle(0, 0, ResetW, ResetH))
                    End Using
                    newBmp.MakeTransparent(Color.FromArgb(255, 255, 0, 128))
                    bmpImg = DrawToBaseImage(newBmp, bmpImg, -1, DefaultResetW, DefaultResetH)
                End If
                Return bmpImg
            End If
        End If
        Return bmpImg
    End Function

    'Private Sub EditAnimations_KeyUp(ByVal chkSender As CheckBox, ByVal e As System.EventArgs)
    '    Dim atemp As Animation = AListClone(chkSender.Tag) 'panRadioPreviewControl
    '    Dim index As Short = chkSender.Tag
    '    pPreview(index).Image = GetPreviewImg(BList(index).Item(0), 0)
    '    atemp.FrameIndex = 0
    '    AListClone(index) = atemp
    '    If chkSender.Checked = False Then
    '        tmrPreview(index).Enabled = False
    '    Else
    '        tmrPreview(chkSender.Tag).Interval = atemp.FrameTimer(0)
    '        tmrPreview(chkSender.Tag).Enabled = True
    '    End If
    '    Dim TrueC As Byte = 0
    '    Dim FalseC As Byte = 0
    '    For i = 0 To panAnimations.Count - 1
    '        If panAnimations(i).Controls.ContainsKey("chk" & i) Then
    '            Dim tempCB As CheckBox = panAnimations(i).Controls("chk" & i)
    '            If tempCB.Checked = True Then TrueC += 1 Else FalseC += 1
    '        End If
    '    Next
    '    If chkSender.FindForm.Controls.ContainsKey("panRadioPreviewControl") Then
    '        Dim tempPan As Panel = chkSender.FindForm.Controls("panRadioPreviewControl")
    '        Dim chkIndex As Byte = 0
    '        If FalseC = panAnimations.Count Then
    '            chkIndex = 2
    '        ElseIf TrueC <> panAnimations.Count Then
    '            chkIndex = 1
    '        End If
    '        If tempPan.Controls.ContainsKey("radCheckbox" & chkIndex) Then
    '            Dim tempRB As RadioButton = tempPan.Controls("radCheckbox" & chkIndex)
    '            tempRB.Checked = True
    '        End If
    '    End If
    'End Sub

    Private Sub mnuTextbox_FrameTimeChanged(ByVal txtSender As TextBox, ByVal e As System.Windows.Forms.KeyEventArgs)
        If txtSender.Text = "" Then txtSender.Text = "0"
        If txtSender.Name.StartsWith("tPreview") Then
            Dim index As Short = txtSender.Name.Substring("tPreview".Length, txtSender.Name.Length - "tPreview".Length)
            If CInt(txtSender.Text) > 32000 Then txtSender.Text = "32000"
            frmAnimations_PreviewRepeatInterval(index) = CInt(txtSender.Text)
            Exit Sub
        End If
        If Val(txtSender.Text) > 32000 Then txtSender.Text = 32000 Else If Val(txtSender.Text) < 1 Then txtSender.Text = 1
        If AListClone(txtSender.Parent.Tag).FrameTimer(txtSender.Tag) <> CShort(txtSender.Text) Then
            AListClone(txtSender.Parent.Tag).FrameTimer(txtSender.Tag) = CShort(txtSender.Text)
            frmAnimations_ChangesMade = True
        End If
    End Sub

    Private Sub tmrPreview_Tick(ByVal tmrSender As Timer, ByVal e As System.EventArgs)
        Dim tempAnimation As Animation = AListClone(tmrSender.Tag)
        If tempAnimation.FrameIndex = -1 Then
            tempAnimation.FrameIndex = 0
            'pPreview(tmrSender.Tag).Image = GetPreviewImg(BList(tmrSender.Tag).Item(tempAnimation.FrameIndex), 0)
            If frmAnimations_PreviewRepeatInterval(tmrSender.Tag) > 0 Then
                tmrSender.Interval = frmAnimations_PreviewRepeatInterval(tmrSender.Tag)
                AListClone(tmrSender.Tag) = tempAnimation
                If tmrSender.Interval > 200 Then pPreview(tmrSender.Tag).Image = AlphaBlendImg(GetPreviewImg(BList(tmrSender.Tag).Item(tempAnimation.FrameIndex), 0), 55)
                Exit Sub
            End If
        End If
        pPreview(tmrSender.Tag).Image = GetPreviewImg(BList(tmrSender.Tag).Item(tempAnimation.FrameIndex), 0)
        tempAnimation.FrameIndex += 1
        If tempAnimation.FrameIndex = tempAnimation.FrameTimer.Count Then
            tempAnimation.FrameIndex = -1
            If tempAnimation.InfiniteRepeat = True Or tempAnimation.RepeatCounter < tempAnimation.Repeats Or 1 = 1 Then
                If tempAnimation.RepeatCounter < tempAnimation.Repeats Then tempAnimation.RepeatCounter += 1
                tmrSender.Interval = tempAnimation.FrameTimer(0) 'frmAnimations_PreviewRepeatInterval(tmrSender.Tag)
            End If
        Else
            tmrSender.Interval = tempAnimation.FrameTimer(tempAnimation.FrameIndex - 1)
        End If
        AListClone(tmrSender.Tag) = tempAnimation
    End Sub

    Private Sub PopulateListbox(ByVal lstBox As ListBox)
        arrIndexes = New List(Of Short)
        lstBox.Items.Clear()
        Dim schTxt As String = txtSearchText.ToUpper
        For i = 0 To AnimationsList.Count - 1
            Dim sLine As String = AnimationsList(i).Substring(0, AnimationsList(i).LastIndexOf("."))
            If sLine.ToUpper.Contains(schTxt) Then
                lstBox.Items.Add(sLine)
                arrIndexes.Add(i)
            End If
        Next
    End Sub

    Private Sub frmAnimations_Textbox_KeyUp(ByVal txtSender As TextBox, ByVal e As KeyEventArgs)
        If txtSender.Name = "txtSearchAnimations" Then
            txtSearchText = txtSender.Text.Trim
            PopulateListbox(lstBox)
            If e.KeyCode = Keys.Enter Then
                If lstBox.Items.Count = 1 Then
                    mnuEdit_LoadAnimation(arrIndexes(0))
                    lstBox.SelectedIndex = 0
                End If
            End If
        End If
    End Sub

    Private Sub frmAnimations_ListBox_SelectedIndexChanged(ByVal lstSender As ListBox, ByVal e As System.EventArgs)
        If lstSender.SelectedIndex = -1 Or lstSender.SelectedIndex = frmAnimations_PrevIndex Then Exit Sub
        mnuEdit_LoadAnimation(arrIndexes(lstSender.SelectedIndex))
        frmAnimations_PrevIndex = lstSender.SelectedIndex
    End Sub

    Private Sub CreateAListClone()
        ReDim AListClone(AList.Count - 1)
        For i = 0 To AList.Count - 1
            AListClone(i).FrameIndex = AList(i).FrameIndex
            AListClone(i).InfiniteRepeat = AList(i).InfiniteRepeat
            AListClone(i).RepeatCounter = AList(i).RepeatCounter
            AListClone(i).FrameTimer = New List(Of Short)
            For i2 = 0 To AList(i).FrameTimer.Count - 1
                AListClone(i).FrameTimer.Add(AList(i).FrameTimer(i2))
            Next
        Next
    End Sub

    Private Sub mnuEdit_LoadAnimation(ByVal index As Short)
        If Not animationTimer Is Nothing Then
            animationTimer.Enabled = False
            animationTimer.Dispose()
        End If
        LoadAnimation(fnAnimations & "\" & AnimationsList(index))
        SetSmileyIndex(index)
        CreateAListClone()
        LoadPanels()
        If panAnimations(0).FindForm.Controls.ContainsKey("txtRenameAnimation") Then
            Dim tempTB As TextBox = panAnimations(0).FindForm.Controls("txtRenameAnimation")
            tempTB.Text = AnimationsList(SIndex)
        End If
    End Sub

#Region "Map Selection"
    Private Sub mnuGame_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim mnuItem As ToolStripMenuItem = sender
        If mnuItem.Checked = True Then mnuItem.Checked = False Else mnuItem.Checked = True
        Dim mnuItemTemp As ToolStripMenuItem = Nothing
        Dim blnChanged As Boolean = False
        For i = 0 To mnuItem.GetCurrentParent.Items.Count - 1
            mnuItemTemp = mnuItem.GetCurrentParent.Items(i)
            If mnuItem.Text <> mnuItemTemp.Text Then
                If mnuItemTemp.Checked = True Then
                    mnuItemTemp.Checked = False
                    PrevioussMenuItem = i
                    blnChanged = True
                    Exit For
                End If
            End If
        Next
        If mnuItem.Text = "Custom" Then
            PrevioussMenuItem = -1
            blnChanged = True
            mnuItem.Checked = True
        End If
        If blnChanged = False Then
            mnuItem.Checked = True
        Else
            If blnFirstMove = True And blnGameOver = False Then
                If MessageBox.Show("Abandon current game?", "uh do it?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                    mnuItemTemp.Checked = True
                    mnuItem.Checked = False
                    Exit Sub
                End If
            End If
            SetDifficulty(mnuItem.Text, Nothing, GameType, TileW, TileH, clr, True)
            If mnuItem.Text = "0 - Standard" Then NewGame(Nothing, W, H, MineCount, 0, TileW, TileH, clr, True)
            If mnuItem.Text = "1 - Hexagonal" Then NewGame(Nothing, W, H, MineCount, 1, TileW, TileH, clr, True)
            If mnuItem.Text = "Custom" Then
                CustomMap()
                If blnCustomMap = False Then
                    mnuItemTemp.Checked = True
                    If mnuItem.Text <> mnuItemTemp.Text Then mnuItem.Checked = False
                Else
                    blnCustomMap = False
                End If
            End If

        End If
    End Sub

    Private Function SetDifficulty(ByRef dif As String, Optional ByVal prnt As Control = Nothing, Optional ByVal gt As Byte = 0, Optional ByRef tw As Short = 16, Optional ByRef th As Short = 16, Optional ByRef tClr As Color = Nothing, Optional ByRef bClear As Boolean = True)
        If dif = "Beginner" Then
            NewGame(prnt, 9, 9, 10, gt, tw, th, tClr, bClear)
            Return True
        End If
        If dif = "Intermediate" Then
            NewGame(prnt, 16, 16, 40, gt, tw, th, tClr, bClear)
            Return True
        End If
        If dif = "Expert" Then
            NewGame(prnt, 30, 16, 100, gt, tw, th, tClr, bClear)
            Return True
        End If
        Return False
    End Function

    Private Sub CustomMap()
        Dim frmCustom As New Form()
        frmCustom.FormBorderStyle = FormBorderStyle.FixedToolWindow
        frmCustom.Text = "Custom Map"
        frmCustom.MaximizeBox = False
        frmCustom.MinimizeBox = False
        frmCustom.ControlBox = True
        Dim pX As Short = 5
        Dim newButton As Button
        For i = 0 To 3
            Dim newLabel As New Label()
            Dim newTextbox As New TextBox()
            If i = 0 Or i = 3 Then
                newButton = New Button()
            Else
                newButton = Nothing
            End If
            newLabel.Parent = frmCustom
            newTextbox.Parent = frmCustom
            If i = 0 Then
                newLabel.Text = "Width:"
                newTextbox.Name = "txtWidth"
            ElseIf i = 1 Then
                newLabel.Text = "Height:"
                newTextbox.Name = "txtHeight"
            ElseIf i = 2 Then
                newLabel.Text = "Mines:"
                newTextbox.Name = "txtMines"
            ElseIf i = 3 Then
                newLabel.Text = "Board Type:"
                newTextbox.Name = "txtBoardType"
                newTextbox.Text = "0"
            End If
            newLabel.AutoSize = True
            'If i <> 0 Then newLabel.Left = newTextbox.Right + 5 Else newLabel.Left = 5
            newLabel.Left = pX
            newTextbox.Left = newLabel.Right + 5
            newTextbox.Width = 30
            newLabel.Top = 10
            newTextbox.Top = 10
            newTextbox.TabIndex = i
            pX = newTextbox.Right + 5

            If Not newButton Is Nothing Then
                newButton.Height = 20
                newButton.Width = 80
                newButton.Top = newLabel.Bottom + 15
                If i = 0 Then
                    newButton.Text = "Cancel"
                    newButton.Left = newLabel.Left
                    newButton.TabIndex = 4
                    frmCustom.CancelButton = newButton
                Else
                    frmCustom.Width = pX + 5
                    frmCustom.Height = newButton.Bottom + 30
                    frmCustom.AcceptButton = newButton
                    newButton.Text = "Create Map"
                    newButton.Left = frmCustom.Width - newButton.Width - 10
                    newButton.TabIndex = 5
                End If
                AddHandler newButton.Click, AddressOf CustomMapClick
                frmCustom.Controls.Add(newButton)
            End If
            AddHandler newTextbox.KeyDown, AddressOf mnuTextbox_NumbersOnly
            frmCustom.Controls.Add(newLabel)
            frmCustom.Controls.Add(newTextbox)
        Next
        frmCustom.StartPosition = FormStartPosition.CenterScreen
        frmCustom.ShowDialog()
    End Sub

    Private Sub CustomMapClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btn As Button = sender
        If btn.Text = "Cancel" Then
            Dim mnuMainTemp As ToolStripMenuItem = MainMenu.Items(0)
            Dim mnuItemTemp As ToolStripMenuItem = mnuMainTemp.DropDownItems(mnuMainTemp.DropDownItems.Count - 1)
            mnuItemTemp.Checked = False
            If PrevioussMenuItem <> -1 Then
                mnuItemTemp = mnuMainTemp.DropDownItems(PrevioussMenuItem)
                mnuItemTemp.Checked = True
            End If
            btn.FindForm.Close()
        ElseIf btn.Text = "Create Map" Then
            Dim newMH As Short = 0
            Dim newMW As Short = 0
            Dim newMC As Short = 0
            Dim newGT As Byte = 0
            Dim controlz As Control
            Dim blnChanged As Boolean = False
            Dim blnOutsideRange As Boolean = False
            For Each controlz In btn.FindForm.Controls
                If controlz.Name.Contains("txt") Then
                    Dim sText As String = controlz.Text.Trim
                    If sText = "" Then
                        MsgBox("Error - " & controlz.Name.Substring(3, controlz.Name.Length - 3) & " field is blank.", MsgBoxStyle.OkOnly, "stop messing up!")
                        controlz.Focus()
                        Exit Sub
                    End If
                    Dim ScreenW As Integer = Math.Floor(Screen.PrimaryScreen.WorkingArea.Width / TileW) - 1
                    Dim screenH As Integer = Math.Floor((Screen.PrimaryScreen.WorkingArea.Height - panGame.Top - 20) / TileH) - 1
                    If controlz.Name.Contains("Height") Or controlz.Name.Contains("Width") Then
                        If CInt(sText) > 100 Then
                            sText = 100
                            blnOutsideRange = True
                        ElseIf CInt(sText) < 8 Then
                            sText = 8
                            blnOutsideRange = True
                        End If
                    End If

                    If controlz.Name.Contains("Width") Then
                        If CInt(sText) > ScreenW Then
                            blnChanged = True
                            sText = ScreenW
                        End If
                        newMW = CInt(sText)
                    End If
                    If controlz.Name.Contains("Height") Then
                        If CInt(sText) > screenH Then
                            blnChanged = True
                            sText = screenH
                        End If

                        newMH = CInt(sText)
                    End If
                    If controlz.Name.Contains("Mines") Then newMC = Math.Max(CInt(sText), 1)
                    If controlz.Name.Contains("BoardType") Then
                        If CInt(sText) > 1 Then sText = 1
                        newGT = CByte(sText)
                    End If

                End If
            Next
            If blnOutsideRange = True Then MsgBox("Map size must be at least 8x8", MsgBoxStyle.OkOnly, "stop messing up!")
            If blnChanged = True Then MsgBox("Size selected is too large for your screen resolution, scaled down to fit.", MsgBoxStyle.OkOnly, "lol")
            blnCustomMap = True
            NewGame(Nothing, newMW, newMH, newMC, newGT, TileW, TileH, clr, True)
            btn.FindForm.Close()
        End If
    End Sub

    Private Function txtNumbersOnly(ByVal sender As TextBox, ByVal kc As Integer, Optional ByRef blnShift As Boolean = False) As Boolean
        If blnShift = True Or ((kc < 48 Or kc > 57) And kc <> Keys.Back And kc <> Keys.Left And kc <> Keys.Right And kc <> Keys.Delete) Then Return True
        Return False
    End Function

    Private Sub mnuTextbox_NumbersOnly(ByVal sender As TextBox, ByVal e As System.Windows.Forms.KeyEventArgs)
        e.SuppressKeyPress = txtNumbersOnly(sender, e.KeyCode, e.Shift)
    End Sub
#End Region

#End Region

#Region "Animation Events"
    Private objAnimation As Animation = Nothing
    Private objBMP As List(Of Bitmap)
    Private AList As List(Of Animation)
    Private BList As List(Of List(Of Bitmap))
    Private AIndex As Byte = 0
    Private ResetDown As Byte = 0

    Private SmileyUp As List(Of Bitmap)
    Private SmileyDown As List(Of Bitmap)
    Private SmileyGameover As List(Of Bitmap)
    Private SmileyVictory As List(Of Bitmap)

    Private SmileyUpA As Animation
    Private SmileyDownA As Animation
    Private SmileyGameoverA As Animation
    Private SmileyVictoryA As Animation

    Private Sub SetSmileyFrame(ByRef iFrame As Byte)
        Dim rect As New Rectangle(0, 0, ResetW, ResetH)
        Dim bmpReset As Bitmap = MineStates(ResetIndex + ResetDown).Clone(rect, Imaging.PixelFormat.DontCare)
        pReset.Image = DrawToBaseImage(bmpReset, objBMP.Item(iFrame), -1, DefaultResetW, DefaultResetH, ResetDown)
    End Sub

    Private Sub SetSmileyIndex(ByRef i As Byte)
        If i >= AnimationsList.Count Then i = AnimationsList.Count - 1
        SIndex = i
    End Sub

    Private Sub AnimateSequence(ByVal i As Byte)
        If Not animationTimer Is Nothing Then
            animationTimer.Enabled = False
            animationTimer.Dispose()
        End If
        objAnimation = AList(i)
        objBMP = BList(i)
        AIndex = i
        animationTimer = New Timer
        AddHandler animationTimer.Tick, AddressOf tmr_Tick
        animationTimer.Interval = 1
        animationTimer.Enabled = True
    End Sub

    Private Sub tmr_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SetSmileyFrame(objAnimation.FrameIndex)
        objAnimation.FrameIndex += 1
        If objAnimation.FrameIndex = objAnimation.FrameTimer.Count Then
            objAnimation.FrameIndex = 0
            If objAnimation.InfiniteRepeat = True Or objAnimation.RepeatCounter < objAnimation.Repeats Then
                If objAnimation.RepeatCounter < objAnimation.Repeats Then objAnimation.RepeatCounter += 1
                animationTimer.Interval = objAnimation.FrameTimer(objAnimation.FrameTimer.Count - 1)
            Else
                animationTimer.Enabled = False
                animationTimer.Dispose()
            End If
        Else
            animationTimer.Interval = objAnimation.FrameTimer(objAnimation.FrameIndex - 1)
        End If
    End Sub

#End Region

#Region "Mouse Events"
    Private Sub ClearHightlights()
        For i = 0 To lstHightlight.Count - 1
            SetTileIndex(lstHightlight(i).X, lstHightlight(i).Y, ButtonUpIndex)
        Next
        lstHightlight = New List(Of Point)
    End Sub

    Private Sub CheckHiglightArea(ByRef x As Byte, ByRef y As Byte)
        ClearHightlights()
        Dim ext As Short = x Mod 2
        Dim iCount As Short = 0
        Dim lstFlags As New List(Of Point)
        Dim lstSpaces As New List(Of Point)
        For r = x - 1 To x + 1
            For c = y - 1 To y + 1
                If (r = x And c = y) Or r < 0 Or c < 0 Or r = W Or c = H Then Continue For
                If GameType = 1 Then If ext = 1 Then If (c = y - 1 And r <> x) Then Continue For Else  Else If (c = y + 1 And r <> x) Then Continue For
                If Tiles(r, c, 0) = FlagIndex Then lstFlags.Add(New Point(r, c)) Else If (Tiles(r, c, 0) = ButtonUpIndex Or Tiles(r, c, 0) = QuestionIndex) Then lstSpaces.Add(New Point(r, c))
            Next
        Next
        If lstFlags.Count = Tiles(x, y, 0) Then
            For i = 0 To lstFlags.Count - 1
                If Tiles(lstFlags(i).X, lstFlags(i).Y, 1) <> MineIndex Then
                    SetTileIndex(lstFlags(i).X, lstFlags(i).Y, MineIndex)
                    GameOver(lstFlags(i).X, lstFlags(i).Y)
                    Exit Sub
                End If
            Next 'If flags correctly placed
            For i = 0 To lstSpaces.Count - 1
                CalculateResult(lstSpaces(i).X, lstSpaces(i).Y)
            Next
        End If
        lstFlags = Nothing
        lstSpaces = Nothing
    End Sub

    Private Sub HighlightSpaces(ByRef x As Byte, ByRef y As Byte)
        lstHightlight = New List(Of Point)
        Dim ext As Short = x Mod 2
        Dim iCount As Short = 0
        For r = x - 1 To x + 1
            For c = y - 1 To y + 1
                If (r = x And c = y) Or r < 0 Or c < 0 Or r = W Or c = H Then Continue For
                If GameType = 1 Then If ext = 1 Then If (c = y - 1 And r <> x) Then Continue For Else  Else If (c = y + 1 And r <> x) Then Continue For
                If Tiles(r, c, 0) = FlagIndex Then iCount += 1 Else If Tiles(r, c, 0) = ButtonUpIndex Then lstHightlight.Add(New Point(r, c))
            Next
        Next
        For i = 0 To lstHightlight.Count - 1
            SetTileIndex(lstHightlight(i).X, lstHightlight(i).Y, 0)
        Next
        blnSelected = True
    End Sub

    Public Sub pMineField_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        pMF.MouseX = e.X
        pMF.MouseY = e.Y
        If pMF.MouseX < 0 Then
            pMF.blnMouseOffPicture = True
            pMF.MouseX = 0
        End If
        If pMF.MouseY < 0 Then
            pMF.blnMouseOffPicture = True
            pMF.MouseY = 0
        End If
        pMF.MouseX = Math.Floor(pMF.MouseX / TileW)
        Dim ext As Short = CheckExt(pMF.MouseX)
        pMF.MouseY = Math.Floor((pMF.MouseY - ext) / TileH)

        If pMF.MouseX >= W Then
            pMF.MouseX = W - 1
            pMF.blnMouseOffPicture = True
        End If
        If pMF.MouseY >= H Then
            pMF.MouseY = H - 1
            pMF.blnMouseOffPicture = True
        End If
        If e.X > -1 And e.X < W * TileW And e.Y > -1 And e.Y < (H * TileH) + ext + 2 Then
            pMF.blnMouseOffPicture = False
        End If
        If pMF.MouseX < 0 Then pMF.MouseX = 0
        If pMF.MouseY < 0 Then pMF.MouseY = 0

        If blnSelected = True Then
            If pMF.MouseX <> pMF.MouseDownX Or pMF.MouseY <> pMF.MouseDownY Then
                ClearHightlights()
                If Tiles(pMF.MouseX, pMF.MouseY, 0) > 0 And Tiles(pMF.MouseX, pMF.MouseY, 0) < 9 Then
                    pMF.MouseDownX = pMF.MouseX
                    pMF.MouseDownY = pMF.MouseY
                    HighlightSpaces(pMF.MouseX, pMF.MouseY)
                End If
                pMinefield.Image = img
            End If
        End If
    End Sub

    Public Sub pMineField_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If blnGameOver = True Then Exit Sub
        If e.Button = Windows.Forms.MouseButtons.Left Then pMF.blnLMDown = True
        If e.Button = Windows.Forms.MouseButtons.Right Then pMF.blnRMDown = True
        If pMF.blnLMDown = True And pMF.blnRMDown = True Then
            If Tiles(pMF.MouseX, pMF.MouseY, 0) > 0 And Tiles(pMF.MouseX, pMF.MouseY, 0) < 9 Then
                HighlightSpaces(pMF.MouseX, pMF.MouseY)
                pMinefield.Image = img
                Exit Sub
            End If
        End If
        If e.Button = Windows.Forms.MouseButtons.Left Then
            AnimateSequence(1)
            If Tiles(pMF.MouseX, pMF.MouseY, 0) = ButtonUpIndex Then
                SetTileIndex(pMF.MouseX, pMF.MouseY, 0)
                pMF.MouseDownX = pMF.MouseX
                pMF.MouseDownY = pMF.MouseY
            End If
            DrawToFrame(pMF.MouseX, pMF.MouseY)
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
            If blnFirstMove = False Then Exit Sub
            Select Case Tiles(pMF.MouseX, pMF.MouseY, 0)
                Case ButtonUpIndex
                    SetTileIndex(pMF.MouseX, pMF.MouseY, FlagIndex)
                    UpdateMinesLeft(-1)
                Case FlagIndex
                    If blnEnableQuestion = True Then SetTileIndex(pMF.MouseX, pMF.MouseY, QuestionIndex) Else SetTileIndex(pMF.MouseX, pMF.MouseY, ButtonUpIndex)
                    UpdateMinesLeft(1)
                Case QuestionIndex
                    SetTileIndex(pMF.MouseX, pMF.MouseY, ButtonUpIndex)
            End Select
        End If
        pMinefield.Image = img
    End Sub

    Public Sub pMineField_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If blnGameOver = True Then Exit Sub
        If e.Button = Windows.Forms.MouseButtons.Left Then pMF.blnLMDown = False
        If e.Button = Windows.Forms.MouseButtons.Right Then pMF.blnRMDown = False
        If pMF.blnLMDown = True And pMF.blnRMDown = True Then
        Else
            If blnSelected = True Then
                blnSelected = False
                If pMF.blnMouseOffPicture = False Then
                    CheckHiglightArea(pMF.MouseX, pMF.MouseY)
                    Exit Sub
                End If
            End If
        End If
        If e.Button = Windows.Forms.MouseButtons.Left Then
            If pMF.blnMouseOffPicture = True Then
                If Tiles(pMF.MouseDownX, pMF.MouseDownY, 0) = 0 Then SetTileIndex(pMF.MouseDownX, pMF.MouseDownY, ButtonUpIndex)

            Else
                If (pMF.MouseDownX <> pMF.MouseX Or pMF.MouseDownY <> pMF.MouseY) Then
                    If blnMouseDownMatch = True Then
                        If Tiles(pMF.MouseDownX, pMF.MouseDownY, 0) = 0 Then SetTileIndex(pMF.MouseDownX, pMF.MouseDownY, ButtonUpIndex)
                    Else
                        CalculateResult(pMF.MouseDownX, pMF.MouseDownY)
                        If blnGameOver = True Then Exit Sub
                    End If
                Else
                    If Tiles(pMF.MouseX, pMF.MouseY, 0) = 0 Then
                        CalculateResult(pMF.MouseX, pMF.MouseY)
                        If blnGameOver = True Then Exit Sub
                    End If
                End If
            End If
        End If
        pMinefield.Image = img
        AnimateSequence(0)
    End Sub

    Private Sub pReset_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
        ResetDown = 1
        Dim rect As New Rectangle(0, 0, ResetW, ResetH)
        Dim bmpReset As Bitmap = MineStates(ResetIndex + ResetDown).Clone(rect, Imaging.PixelFormat.DontCare)
        Dim index As Short = objAnimation.FrameIndex - 1
        If index < 0 Then index = objAnimation.FrameTimer.Count - 1
        pReset.Image = DrawToBaseImage(bmpReset, objBMP.Item(index), -1, DefaultResetW, DefaultResetH, 1)
    End Sub

    Private Sub pReset_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
        ResetDown = 0
        If e.X < 0 Or e.X > pReset.Width Or e.Y < 0 Or e.Y > pReset.Height Then
            Dim rect As New Rectangle(0, 0, ResetW, ResetH)
            Dim bmpReset As Bitmap = MineStates(ResetIndex + ResetDown).Clone(rect, Imaging.PixelFormat.DontCare)
            Dim index As Short = objAnimation.FrameIndex - 1
            If index < 0 Then index = objAnimation.FrameTimer.Count - 1
            pReset.Image = DrawToBaseImage(bmpReset, objBMP.Item(index), -1, DefaultResetW, DefaultResetH)
        Else
            NewGame(Nothing, W, H, MineCount, GameType, TileW, TileH, clr, True)
        End If
    End Sub
#End Region

#Region "Graphics / Colors"
    Public Function LoadGraphics() As Boolean
        Try
            Dim clrDarker As Color
            Dim clrBorder As Color
            LoadAnimation(fnAnimations & "\" & AnimationsList(SIndex))

            MineStates = New List(Of Bitmap)
            MineStates.Add(Image.FromFile(fnContent & "\ButtonDown.png"))
            Dim sourceRect As New Rectangle(0, 0, DefaultTileW, DefaultTileH)
            If clr <> Nothing Then
                clrDarker = DarkerColor(clr)
                clrBorder = BorderColor(clr)
                SetTileColors(False, clr, clrDarker)
            End If
            For i = 1 To 9
                Dim newImg As Bitmap
                If i = 9 Then
                    newImg = Image.FromFile(fnContent & "\Mine.png")
                    MineIndex = i
                Else
                    newImg = Image.FromFile(fnContent & "\" & i & ".png")
                End If

                MineStates.Add(DrawToBaseImage(MineStates(0).Clone(sourceRect, Imaging.PixelFormat.DontCare), newImg, i, DefaultTileW, DefaultTileH))
            Next

            MineStates.Add(Image.FromFile(fnContent & "\ButtonUp.png"))
            ButtonUpIndex = MineStates.Count - 1
            If clr <> Nothing Then SetTileColors(True, clr, clrDarker, clrBorder)
            For i = ButtonUpIndex + 1 To ButtonUpIndex + 2
                Dim newImg As Bitmap
                If i = ButtonUpIndex + 1 Then
                    newImg = Image.FromFile(fnContent & "\Flag.png")
                    FlagIndex = i
                Else
                    newImg = Image.FromFile(fnContent & "\QuesitonMark.png")
                    QuestionIndex = i
                End If
                MineStates.Add(DrawToBaseImage(MineStates(ButtonUpIndex).Clone(sourceRect, Imaging.PixelFormat.DontCare), newImg, i))
            Next

            Dim bmpCross As New Bitmap(TileW, TileH)
            For r = 0 To bmpCross.Width - 1
                For c = 0 To bmpCross.Height - 1
                    bmpCross.SetPixel(r, c, Color.FromArgb(255, 255, 0, 128))
                Next
            Next
            Using gx As Graphics = Graphics.FromImage(bmpCross)
                For i = 0 To 1
                    gx.DrawLine(Pens.Red, i, 0, bmpCross.Width, bmpCross.Height - i)
                    gx.DrawLine(Pens.Red, 0, i, bmpCross.Width - i, bmpCross.Height)

                    gx.DrawLine(Pens.Red, bmpCross.Width - i, 0, 0, bmpCross.Height - i)
                    gx.DrawLine(Pens.Red, bmpCross.Width, i, i, bmpCross.Height)
                Next

            End Using
            bmpCross.MakeTransparent(Color.FromArgb(255, 255, 0, 128))
            CrossIndex = MineStates.Count
            MineStates.Add(bmpCross)

            ResetIndex = MineStates.Count 'Load smiley animations
            Dim newBMP As New Bitmap(ResetW, ResetH)
            Using gx As Graphics = Graphics.FromImage(newBMP)
                gx.DrawImage(MineStates(ButtonUpIndex).Clone(sourceRect, Imaging.PixelFormat.DontCare), 0, 0, ResetW, ResetH)
                MineStates.Add(newBMP)
            End Using
            ResetIndex2 = MineStates.Count
            newBMP = New Bitmap(ResetW, ResetH)
            Using gx As Graphics = Graphics.FromImage(newBMP)
                gx.DrawImage(MineStates(0).Clone(sourceRect, Imaging.PixelFormat.DontCare), 0, 0, ResetW, ResetH)
                MineStates.Add(newBMP)
            End Using


            Numbers = New List(Of Bitmap)
            Dim NewW As Short = 0
            Dim NewH As Short = 0
            If blnScaleNumbers = True Then
                NewW = (ResetW / DefaultResetW) * DefaultNumberW
                NewH = (ResetH / DefaultResetH) * DefaultNumberH
            End If
            For i = 0 To 10
                Dim sFile As String = fnContent
                If i = 10 Then sFile &= "\NMinus.png" Else sFile &= "\N" & i & ".png"
                If blnScaleNumbers = True Then
                    Numbers.Add(DrawToBaseImage(New Bitmap(NewW, NewH), Image.FromFile(sFile), -1, DefaultNumberW, DefaultNumberH))
                Else
                    Numbers.Add(Image.FromFile(sFile))
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox("Error: Could not load MineStates array.", MsgBoxStyle.OkOnly, "stop messing up!")
            Return False
        End Try
    End Function

    Private Sub LoadAnimation(ByVal sFile As String)
        Dim blnLoaded As Boolean = False
        If System.IO.File.Exists(sFile) Then
            For i = 0 To 3
                AList(i) = New Animation
                BList(i) = New List(Of Bitmap)
            Next
            blnLoaded = CheckAndLoadSequence(sFile.Substring(sFile.LastIndexOf("\") + 1, sFile.Length - sFile.LastIndexOf("\") - 1))
            Dim bmpAnimation As Bitmap = Image.FromFile(sFile)
            For i = 0 To 3
                Dim arrTempBMP As List(Of Bitmap) = BList(i)
                Dim iFrames As Short = 0
                Dim iWidth As Short = Math.Floor((bmpAnimation.Width / SmileyW)) - 1
                For i2 = 0 To iWidth
                    Dim bmpFrame As Bitmap = New Bitmap(SmileyW, SmileyH)
                    Dim rect As New Rectangle(i2 * SmileyW, i * SmileyH, SmileyW, SmileyH)
                    Using gx As Graphics = Graphics.FromImage(bmpFrame)
                        gx.DrawImage(bmpAnimation.Clone(rect, Imaging.PixelFormat.DontCare), 0, 0, SmileyW, SmileyH)
                        iFrames += 1
                        If bmpFrame.GetPixel(0, 0) = Color.FromArgb(255, 255, 255, 255) Then
                            If bmpFrame.GetPixel(bmpFrame.Width - 1, bmpFrame.Height - 1) = Color.FromArgb(255, 255, 255, 255) Then Exit For
                        End If
                        bmpFrame.MakeTransparent(Color.FromArgb(255, 255, 0, 128))
                        arrTempBMP.Add(bmpFrame)
                    End Using
                Next
            Next

        Else
            MsgBox("Error: Could not find file " & sFile, MsgBoxStyle.OkOnly, "stop messing up!")
        End If
        If blnLoaded = False Then
            SmileyW = DefaultSmileyW
            SmileyH = DefaultSmileyH
            For i = 0 To AList.Count - 1
                Dim aniTemp As Animation = AList(i)
                aniTemp.FrameTimer = New List(Of Short)
                aniTemp.InfiniteRepeat = i > 1
                aniTemp.Repeats = 0
                For i2 = 0 To BList(i).Count - 1
                    aniTemp.FrameTimer.Add(100)
                Next
                AList(i) = aniTemp
            Next
        End If
        BList = BList
        AList = AList
    End Sub

    Private Function CheckAndLoadSequence(ByVal sFileName As String) As Boolean
        Dim blnLoaded As Boolean = False
        If System.IO.File.Exists(fnAnimations & "\" & "Sequences.txt") Then
            Dim sr As System.IO.StreamReader = System.IO.File.OpenText(fnAnimations & "\" & "Sequences.txt")
            Do While sr.Peek <> -1 'SmileyExplode.png|[False,0]30,[False,0]30,[False,0]30,30,50,30,[True,0]100,100,
                Dim sLine As String = sr.ReadLine.Trim
                If sLine.Length < 3 Then Continue Do
                If sLine.Substring(0, sLine.IndexOf(":")).ToUpper = sFileName.ToUpper Then
                    blnLoaded = True
                    SmileyW = CByte(sLine.Substring(sLine.IndexOf(":") + 1, sLine.IndexOf(",") - sLine.IndexOf(":") - 1))
                    SmileyH = CByte(sLine.Substring(sLine.IndexOf(",") + 1, sLine.IndexOf("|") - sLine.IndexOf(",") - 1))
                    If SmileyW <> DefaultSmileyW Or SmileyH <> DefaultSmileyH Then
                        Dim xScale As Decimal = SmileyW / DefaultSmileyW
                        Dim yScale As Decimal = SmileyH / DefaultSmileyH
                        ResetW *= xScale
                        ResetH *= yScale
                    End If
                    sLine = sLine.Substring(sLine.IndexOf("|") + 2, sLine.Length - sLine.IndexOf("|") - 2)
                    Dim sRepeats As String = ""
                    For i = 0 To 3
                        Dim aniTemp As Animation = AList(i)
                        Dim sItem As String = ""
                        If i = 3 Then
                            sItem = sLine.Substring(0, sLine.Length)
                        Else
                            sItem = sLine.Substring(0, sLine.IndexOf("["))
                            sLine = sLine.Remove(0, sItem.Length + 1)
                        End If
                        aniTemp.FrameTimer = New List(Of Short)
                        aniTemp.InfiniteRepeat = CBool(sItem.Substring(0, sItem.IndexOf(",")))
                        aniTemp.RepeatCounter = CByte(sItem.Substring(sItem.IndexOf(",") + 1, sItem.IndexOf("]") - sItem.IndexOf(",") - 1))
                        Dim PIndex As Short = sItem.IndexOf("]") + 1
                        Do
                            Dim Index As Short = sItem.IndexOf(",", PIndex)
                            If Index = -1 Then Exit Do
                            aniTemp.FrameTimer.Add(CShort(sItem.Substring(PIndex, Index - PIndex)))
                            PIndex = Index + 1
                        Loop
                        AList(i) = aniTemp
                    Next
                End If
            Loop
            sr.Close()
        End If
        Return blnLoaded
    End Function

    Private Function AlphaBlendImg(ByRef bmpImg As Bitmap, ByRef bAlphaBlend As Byte) As Bitmap
        Using gx As Graphics = Graphics.FromImage(bmpImg)
            gx.FillRectangle(New SolidBrush(Color.FromArgb(bAlphaBlend, 255, 255, 255)), New Rectangle(0, 0, bmpImg.Width, bmpImg.Height))
        End Using
        Return bmpImg
    End Function

    Private Function DrawToBaseImage(ByRef imgBase As Bitmap, ByRef imgObject As Bitmap, Optional ByRef index As Short = -1, Optional ByRef oX As Byte = 0, Optional ByRef oY As Byte = 0, Optional ByRef modif As Byte = 0) As Bitmap
        Using gx As Graphics = Graphics.FromImage(imgBase)
            Dim ScaleX As Double = 1, ScaleY As Double = 1
            If oX <> 0 And (oX <> imgBase.Width Or oY <> imgBase.Height) Then
                ScaleX = imgBase.Width / oX
                ScaleY = imgBase.Height / oY
            End If
            Dim ScaledW As Short = ScaleX * imgObject.Width
            Dim ScaledH As Short = ScaleY * imgObject.Height
            Dim iX As Short = ((imgBase.Width - ScaledW) / 2) + modif
            Dim iY As Short = ((imgBase.Height - ScaledH) / 2) + modif
            If index <> -1 Then If (ScaledH Mod 2) = 1 And index <> MineIndex Then iX += 1
            imgObject.MakeTransparent(Color.FromArgb(255, 255, 0, 128))
            Using gx2 As Graphics = Graphics.FromImage(imgBase)
                gx2.DrawImage(imgObject, iX, iY, ScaledW, ScaledH)
                Return imgBase
            End Using
        End Using
    End Function

    Public Function SetTileColors(ByRef blnType As Boolean, ByRef tClr As Color, ByRef clrDarker As Color, Optional ByRef clrBorder As Color = Nothing)
        Try
            If blnType = False Then
                Dim clrCCenter As Color = MineStates(0).GetPixel(MineStates(0).Width / 2, MineStates(0).Height / 2)
                Dim clrCDarker As Color = MineStates(0).GetPixel(0, 0)
                If clrCCenter <> tClr Then FillColorAllSpaces(MineStates(0), tClr, clrCCenter)
                If clrDarker <> clrCDarker Then FillColorAllSpaces(MineStates(0), clrDarker, clrCDarker)
                Return True
            Else
                Dim clrCCenter As Color = MineStates(ButtonUpIndex).GetPixel(MineStates(ButtonUpIndex).Width / 2, MineStates(ButtonUpIndex).Height / 2)
                Dim clrCDarker As Color = MineStates(ButtonUpIndex).GetPixel(MineStates(ButtonUpIndex).Width - 1, MineStates(ButtonUpIndex).Height - 1)
                Dim clrCBorder As Color = MineStates(ButtonUpIndex).GetPixel(0, 0)
                If clrCCenter <> tClr Then FillColorAllSpaces(MineStates(ButtonUpIndex), tClr, clrCCenter)
                If clrDarker <> clrCDarker Then FillColorAllSpaces(MineStates(ButtonUpIndex), clrDarker, clrCDarker)
                If clrBorder <> clrCBorder Then FillColorAllSpaces(MineStates(ButtonUpIndex), clrBorder, clrCBorder)
                Return True
            End If

        Catch ex As Exception
            MsgBox("Error: SetTileColors failed." & vbNewLine & ex.Message)
            Return False
        End Try
    End Function

    Private Function BorderColor(ByRef tClr As Color) As Color
        Dim dif As Short = 50
        Dim newR As Short = tClr.R + dif
        Dim newB As Short = tClr.B + dif
        Dim newG As Short = tClr.G + dif
        If newR < 0 Then newR = 0 Else If newR > 255 Then newR = 255
        If newB < 0 Then newB = 0 Else If newB > 255 Then newB = 255
        If newG < 0 Then newG = 0 Else If newG > 255 Then newG = 255
        Return Color.FromArgb(255, newR, newG, newB)
    End Function

    Private Function DarkerColor(ByRef tClr As Color) As Color
        Try
            Dim dif As Short = 50
            Dim newR As Short = tClr.R - dif
            Dim newB As Short = tClr.B - dif
            Dim newG As Short = tClr.G - dif
            If newR < 0 Then newR = 0 Else If newR > 255 Then newR = 255
            If newB < 0 Then newB = 0 Else If newB > 255 Then newB = 255
            If newG < 0 Then newG = 0 Else If newG > 255 Then newG = 255
            Return Color.FromArgb(255, newR, newG, newB)
        Catch ex As Exception

        End Try

    End Function

    Private Sub FillColorAllSpaces(ByRef tImg As Bitmap, ByRef newColor As Color, ByRef oldColor As Color, Optional ByRef regionX As Short = 0, Optional ByRef regionY As Short = 0, Optional ByRef regionW As Short = -1, Optional ByRef regionH As Short = -1)
        If regionW = -1 Then regionW = tImg.Width
        If regionH = -1 Then regionH = tImg.Height
        For r = regionX To regionX + regionW - 1
            For c = regionY To regionY + regionH - 1
                If tImg.GetPixel(r, c) = oldColor Then tImg.SetPixel(r, c, newColor)
            Next
        Next
    End Sub
#End Region

End Class