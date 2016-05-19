
Option Explicit On



Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Drawing



Public Class clsSprite

    Public InConversation As Boolean


    'Secondary frames
    Private SFrame As Integer
    Private sFrameLimit As Integer = 9



    'This multiplier will be adjusted according to the speed of the character.
    Public Multiplier As Single = 30 / 30
    'Just a dirty little trick: Adjust the MultiplierNumerator according to Multiplier
    Public MultiplierNumerator As Integer = 30
    'Stores the position (x,y) of the sprite. Sometimes these values will be decimals, so use PointF
    Public Position As PointF
    'Stores the folder in which the Sprite images are located in
    Public Folder As String
    'Stores the current frame (1 2 or 3) of the sprite
    Public Frame As Integer

    'Stores the number of Directions each sprite can have
    Public Const NumberOfDirections As Integer = 4
    'Stores the number of Frames each sprite can have
    Public Const NumberOfFrames As Integer = 3

    'Holds all the images for the sprite
    Public SpriteImage(NumberOfDirections - 1, NumberOfFrames - 1) As Bitmap


    'Stores the current direction of the sprite
    Public Enum Dir
        Up = 0
        Down = 1
        Left = 2
        Right = 3
    End Enum
    'Create a copy for our sprite
    Public Direction As Dir

    Public SpriteName As String

    Public InMotion As Boolean

    Public ReadOnly Property TilePos() As Point
        Get
            Return New Point(Math.Floor(Position.X / GameClass.TileWidth), Math.Floor(Position.Y / GameClass.TileHeight))
        End Get
    End Property



    'Constructor. 
    Public Sub New(ByVal NameOfSprite As String)
        'Store the SpriteName (the sprites are stored in bin\Sprites\SpriteName)
        SpriteName = NameOfSprite
        Dim CurrentDirection As Integer
        Dim CurrentFrame As Integer
        'Loop through every Direction
        For CurrentDirection = 0 To NumberOfDirections - 1
            For CurrentFrame = 0 To NumberOfFrames - 1
                'Set each bitmap (all 12)
                SpriteImage(CurrentDirection, CurrentFrame) = New Bitmap("Sprites\" & NameOfSprite & "\" & ReturnDirection(CurrentDirection) & CurrentFrame + 1 & ".GIF") 'These aren't 2 different lines, the page isn't wide enough to fit the entire line.
                SpriteImage(CurrentDirection, CurrentFrame).MakeTransparent(Color.Lime)
            Next
        Next
    End Sub


    'ReturnDirection returns a String from an Integer value provided by the argument
    Public Function ReturnDirection(ByVal CurrentDirection As Integer) As String
        If CurrentDirection = Dir.Up Then Return "up"
        If CurrentDirection = Dir.Down Then Return "down"
        If CurrentDirection = Dir.Left Then Return "left"
        If CurrentDirection = Dir.Right Then Return "right"
    End Function

    Public Sub MoveUp()

        If GameClass.Physics.Allow(Dir.Up) Then

            'Set his direction
            Direction = Dir.Up
            'Animate the sprite!
            Me.Animate()
            GameClass.GameForm.Invalidate()


            'If the sprite is currently not moving Then move him up
            If Not InMotion Then
                'Now he's in motion.
                InMotion = True
                'Loop Variable
                Dim x As Single
                Me.Animate()

                'Set his direction
                Direction = Dir.Up
                'Loop 30 times controlled by the multiplier
                For x = 1 To GameClass.TileHeight * Multiplier
                    'Move him one unit up, 30 times controlled by the multiplier
                    Position.Y -= 1 / Multiplier
                    'Animate the sprite!
                    Me.Animate()
                    'Let your application do other events (ex: Check for keypress and refresh the form)
                    'in this loop.
                    Application.DoEvents()
                    'Refresh the form
                    GameClass.GameForm.Invalidate()
                Next
                'Round his decimal place
                Position.Y = Math.Round(Position.Y, 0)
                'Now he's done moving. 
                InMotion = False
            End If
        End If

    End Sub

    Public Sub MoveDown()

        'Set his direction
        Direction = Dir.Down
        'Animate the sprite!
        Me.Animate()
        GameClass.GameForm.Invalidate()

        If GameClass.Physics.Allow(Dir.Down) Then

            'If the sprite is currently not moving Then move him up
            If Not InMotion Then
                'Now he's in motion.
                InMotion = True
                'Loop Variable
                Dim x As Single
                Me.Animate()
                'Set his direction
                Direction = Dir.Down
                'Loop 30 times controlled by the multiplier
                For x = 1 To GameClass.TileHeight * Multiplier
                    'Move him one unit up, 30 times controlled by the multiplier
                    Position.Y += 1 / Multiplier
                    'Animate the sprite!
                    Me.Animate()
                    'Let your application do other events (ex: Check for keypress and refresh the form)
                    'in this loop.
                    Application.DoEvents()
                    'Refresh the form
                    GameClass.GameForm.Invalidate()
                Next
                'Round his decimal place
                Position.Y = Math.Round(Position.Y, 0)
                'Now he's done moving. 
                InMotion = False
            End If
        End If

    End Sub

    Public Sub MoveLeft()


        'Set his direction
        Direction = Dir.Left
        'Animate the sprite!
        Me.Animate()
        GameClass.GameForm.Invalidate()
        If GameClass.Physics.Allow(Dir.Left) Then

            'If the sprite is currently not moving Then move him up
            If Not InMotion Then
                'Now he's in motion.
                InMotion = True
                'Loop Variable
                Dim x As Single
                Me.Animate()
                'Set his direction
                Direction = Dir.Left
                'Loop 30 times controlled by the multiplier
                For x = 1 To GameClass.TileWidth * Multiplier
                    'Move him one unit up, 30 times controlled by the multiplier
                    Position.X -= 1 / Multiplier
                    'Animate the sprite!
                    Me.Animate()
                    'Let your application do other events (ex: Check for keypress and refresh the form)
                    'in this loop.
                    Application.DoEvents()
                    'Refresh the form
                    GameClass.GameForm.Invalidate()
                Next
                'Round his decimal place
                Position.X = Math.Round(Position.X, 0)
                'Now he's done moving. 
                InMotion = False
            End If
        End If

    End Sub

    Public Sub MoveRight()


        'Set his direction
        Direction = Dir.Right
        'Animate the sprite!
        Me.Animate()
        GameClass.GameForm.Invalidate()

        If GameClass.Physics.Allow(Dir.Right) Then

            'If the sprite is currently not moving Then move him up
            If Not InMotion Then
                'Now he's in motion.
                InMotion = True
                'Loop Variable
                Dim x As Single
                Me.Animate()
                'Set his direction
                Direction = Dir.Right
                'Loop 30 times controlled by the multiplier
                For x = 1 To GameClass.TileWidth * Multiplier
                    'Move him one unit up, 30 times controlled by the multiplier
                    Position.X += 1 / Multiplier
                    'Animate the sprite!
                    Me.Animate()
                    'Let your application do other events (ex: Check for keypress and refresh the form)
                    'in this loop.
                    Application.DoEvents()
                    'Refresh the form
                    GameClass.GameForm.Invalidate()
                Next
                'Round his decimal place
                Position.X = Math.Round(Position.X, 0)
                'Now he's done moving. 
                InMotion = False
            End If
        End If

    End Sub


    Public Sub Animate()
        SFrame += 1

        Select Case MultiplierNumerator
            Case Is >= 30
                SFrameLimit = 9
            Case Is >= 15
                SFrameLimit = 6
            Case Is >= 6
                SFrameLimit = 3
        End Select

        If SFrame > SFrameLimit Then SFrame = SFrameLimit

        If SFrame = SFrameLimit Then
            SFrame = 0
            ' Make him move
            Frame += 1
            ' Don't let it go too high!
            If Frame = 3 Then Frame = 0
        End If

    End Sub




End Class

Public Class GameClass
    Private SetRes As ScreenSettings
    

    Public Shared NPCs As New Collections.Hashtable()
    'This will hold all the items in our NPCs collection. See Alternative to Arrays tutorial for more information.
    Dim Entry As Collections.DictionaryEntry
    Public Shared Level As New clsLevel()

    Public Shared Events As New clsEvent()
    Public Shared alex As New clsSprite("alex")
    Public Shared map As clsMap


    'Declare the physics class
    Public Shared Physics As New clsCollision()

    'The form in which we will draw to.
    Public Shared GameForm As Form

    Public Const TileWidth As Integer = 30
    Public Const TileHeight As Integer = 30

    Public Sub Terminate()
        SetRes.Dispose()
    End Sub
    Public Sub New(ByVal frm As Form)
        'Set the form that we'll draw to
        GameForm = frm
        AddHandler GameForm.Paint, AddressOf Me.Paint
        AddHandler GameForm.KeyDown, AddressOf Me.Keydown
        SetRes = New ScreenSettings(640, 480)


    End Sub







    Public Sub Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)


        
        Dim x As Integer
        Dim y As Integer

        For x = 0 To map.Width
            For y = 0 To map.Height
                'Scroll in the direction opposite to moveGameclassnt

                e.Graphics.DrawImage(Me.map.Tiles(x, y), New Point(x * Me.TileWidth - Me.alex.Position.X + 300, y * Me.TileHeight - Me.alex.Position.Y + 240))
                For Each Entry In NPCs
                    e.Graphics.DrawImage(DirectCast(NPCs(Entry.Key), clsNPC).CurrentImage, New Point((DirectCast(NPCs(Entry.Key), clsNPC).Position.X * 30) - Me.alex.Position.X + 300, DirectCast(NPCs(Entry.Key), clsNPC).Position.Y * 30 - Me.alex.Position.Y + 240))
                Next
            Next
        Next

        'Keep the character centered on the screen. The *real* center is (640/2,480/2) which is (320, 240). Since 320 doesn't divide evenly into 30 (Meaning he won't be on a tile 'if he's placed at an X of 320), I simply rounded down to 300.
        e.Graphics.DrawImage(GameClass.alex.SpriteImage(GameClass.alex.Direction, GameClass.alex.Frame), New Point(300, 240))

        If Not clsText.Text Is Nothing Then
            'Draw the dialogue box
            e.Graphics.DrawImage(New Bitmap("DialogueImage\default.bmp"), New Point(80, 320))
            'Draw the text
            e.Graphics.DrawString(clsText.Text, New Font("Courier New", 12, FontStyle.Bold), Brushes.Black, clsText.Position.X, clsText.Position.Y)
        End If

    End Sub

    Public Sub Keydown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        'Select case is just another way of writing If statements.
        If Not GameClass.alex.InConversation Then
            Select Case e.KeyCode
                Case Keys.Up
                    alex.MoveUp()
                Case Keys.Down
                    alex.MoveDown()
                Case Keys.Left
                    alex.MoveLeft()
                Case Keys.Right
                    alex.MoveRight()
                Case Keys.Delete
                    GameForm.Invalidate()
                
            End Select
        End If
        Select Case e.KeyCode
            Case Keys.Space
                Events.CheckForEvents()
            Case Keys.Escape
                GameForm.Close()
        End Select
        'Display his position. Commented out because this is fullscreen and this isn't needed.
        '   GameClass.GameForm.Text = "Current Position: " & alex.TilePos.ToString
    End Sub
End Class
Public Class clsMap
    'The thing which will read our map
    Dim SR As System.IO.StreamReader
    'Size of the map
    Public Width As Integer
    Public Height As Integer
    'Will store all the tiles
    Public Tiles(,) As Bitmap
    Public Passable(,) As Boolean


    Public Sub New(ByVal MapName As String, ByVal tileset As String)
        ReadMap(MapName, tileset)
    End Sub
    Public Sub ReadMap(ByVal MapName As String, ByVal tileset As String)
        'Read a certain map file
        SR = New System.IO.StreamReader("maps\" & MapName)

        Dim ln As String
        'This will store the width/height of the map, seperated by commas.
        ln = SR.ReadLine()
        'This will store the width and the height individually
        Dim str() As String
        'Split ln from it's commas
        str = ln.Split(",")
        'Store the width and the height
        Width = str(0) - 1
        Height = str(1) - 1

        'Now that we know the height and the width:
        'Redim Tiles according to height and the width of the map
        ReDim Tiles(Width, Height)
        'Stores the current row/column being read

        ReDim Passable(Width, Height)

        Dim TypesOfTiles As Integer
        'Store the next line (which contains the number of types of tiles) into an integer
        TypesOfTiles = SR.ReadLine()
        Dim Walkable(TypesOfTiles) As Boolean

        Dim theLine() As String
        While (Not ln = "#END LEGEND")
            ln = SR.ReadLine
            If Not ln = "#LEGEND" And Not ln = "#END LEGEND" Then
                theLine = ln.Split("=")
                Walkable(theLine(0)) = CBool(theLine(1))
            End If
        End While




        Dim CurrentRow As Integer
        Dim CurrentColumn As Integer

        Dim Line() As String

        For CurrentRow = 0 To Height
            'Read each Line
            ln = SR.ReadLine()
            Line = ln.Split(",")
            For CurrentColumn = 0 To Width
                'Read each character (Go across the file)
                Tiles(CurrentColumn, CurrentRow) = New Bitmap("tiles\" & tileset & "\" & Line(CurrentColumn) & ".jpg")
                Passable(CurrentColumn, CurrentRow) = Walkable(Line(CurrentColumn))
            Next
        Next
    End Sub
End Class



Public Class ScreenSettings

    Const ENUM_CURRENT_SETTINGS As Integer = -1
    Const CDS_UPDATEREGISTRY As Integer = &H1
    Const CDS_TEST As Long = &H2

    Const CCDEVICENAME As Integer = 32
    Const CCFORMNAME As Integer = 32

    Const DISP_CHANGE_SUCCESSFUL As Integer = 0
    Const DISP_CHANGE_RESTART As Integer = 1
    Const DISP_CHANGE_FAILED As Integer = -1

    Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Integer, ByVal iModeNum As Integer, ByRef lpDevMode As DEVMODE) As Integer
    Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (ByRef DEVMODE As DEVMODE, ByVal flags As Integer) As Integer

    <StructLayout(LayoutKind.Sequential)> Public Structure DEVMODE
        <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=CCDEVICENAME)> Public dmDeviceName As String
        Public dmSpecVersion As Short
        Public dmDriverVersion As Short
        Public dmSize As Short
        Public dmDriverExtra As Short
        Public dmFields As Integer
        Public dmOrientation As Short
        Public dmPaperSize As Short
        Public dmPaperLength As Short
        Public dmPaperWidth As Short
        Public dmScale As Short
        Public dmCopies As Short
        Public dmDefaultSource As Short
        Public dmPrintQuality As Short
        Public dmColor As Short
        Public dmDuplex As Short
        Public dmYResolution As Short
        Public dmTTOption As Short
        Public dmCollate As Short
        <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=CCFORMNAME)> Public dmFormName As String
        Public dmUnusedPadding As Short
        Public dmBitsPerPel As Short
        Public dmPelsWidth As Integer
        Public dmPelsHeight As Integer
        Public dmDisplayFlags As Integer
        Public dmDisplayFrequency As Integer
    End Structure

    Dim oldWidth, oldHeight As Integer

    Public Sub New(ByVal width As Integer, ByVal height As Integer)
        oldWidth = Screen.PrimaryScreen.Bounds.Width
        oldHeight = Screen.PrimaryScreen.Bounds.Height

        changeRes(width, height)
    End Sub

    Public Sub Dispose()
        changeRes(oldWidth, oldHeight)
    End Sub

    Private Sub changeRes(ByVal theWidth As Integer, ByVal theHeight As Integer)
        Try
            Dim DevM As DEVMODE

            DevM.dmDeviceName = New [String](New Char(32) {})
            DevM.dmFormName = New [String](New Char(32) {})
            DevM.dmSize = CShort(Marshal.SizeOf(GetType(DEVMODE)))

            If 0 <> EnumDisplaySettings(Nothing, ENUM_CURRENT_SETTINGS, DevM) Then
                Dim lResult As Integer

                DevM.dmPelsWidth = theWidth
                DevM.dmPelsHeight = theHeight

                lResult = ChangeDisplaySettings(DevM, CDS_TEST)

                If lResult = DISP_CHANGE_FAILED Then
                    MsgBox("Display Change Failed.", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Screen Resolution Change Failed")
                Else

                    lResult = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)

                    Select Case lResult
                        Case DISP_CHANGE_RESTART
                            MsgBox("You must restart your computer to apply these changes.", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Screen Resolution Has Changed")
                        Case DISP_CHANGE_SUCCESSFUL

                        Case Else
                            MsgBox("Display Change Failed.", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Screen Resolution Change Failed")
                    End Select
                End If


            End If
        Catch ex As Security.SecurityException
            MsgBox("You do not have authorization to change the resolution")
            System.Environment.Exit(System.Environment.ExitCode)
        End Try
    End Sub

End Class

Public Class clsCollision
    Public Function Allow(ByVal direction As clsSprite.Dir) As Boolean
        Select Case direction
            Case clsSprite.Dir.Up
                'First make sure the tile above him is not the outside of the map
                If GameClass.alex.TilePos.Y = 0 Then Return False
                'Then make sure that the tile above him is passable
                If GameClass.map.Passable(GameClass.alex.TilePos.X, GameClass.alex.TilePos.Y - 1) Then
                    Return True
                Else
                    Return False
                End If
            Case clsSprite.Dir.Down
                'Be sure that he's not on the bottom tile and heading down
                If GameClass.alex.TilePos.Y = GameClass.map.Height Then Return False
                'Then make sure that the tile below him is passable
                If GameClass.map.Passable(GameClass.alex.TilePos.X, GameClass.alex.TilePos.Y + 1) Then
                    Return True
                Else
                    Return False
                End If
            Case clsSprite.Dir.Left
                'Be sure that he's not on the leftmost tile and heading left
                If GameClass.alex.TilePos.X = 0 Then Return False
                'Then make sure that the tile to the left of him is passable
                If GameClass.map.Passable(GameClass.alex.TilePos.X - 1, GameClass.alex.TilePos.Y) Then
                    Return True
                Else
                    Return False
                End If
            Case clsSprite.Dir.Right
                'Be sure that he's not on the rightmost tile and heading right
                If GameClass.alex.TilePos.X = GameClass.map.Width Then Return False
                'Then make sure that the tile to the right of him is passable
                If GameClass.map.Passable(GameClass.alex.TilePos.X + 1, GameClass.alex.TilePos.Y) Then
                    Return True
                Else
                    Return False
                End If
        End Select
    End Function
End Class
Public Class clsNPC
    'Store his position
    Public Position As Point
    'Store his 4 images (Up, Down, Left, Right)
    Public SpriteImage(3) As Bitmap
    'Store his folder where his images are located
    Public Folder As String
    'Our NPC will have a direction, so that he can turn when the character talks to them.
    Public Direction As clsSprite.Dir

    Public Sub New(ByVal ImageFolder As String, ByVal TilePos As Point, ByVal defaultdirection As clsSprite.Dir, ByVal DialogueLoc As String)
        'Store his position
        Position = TilePos
        'Store the variable
        Folder = ImageFolder
        'Set his default direction
        Direction = defaultdirection
        'Load the images
        Load(ImageFolder)
    End Sub

    Public Sub Load(ByVal ImageFolder As String)
        'Up = 0
        'Down = 1
        'Left = 2
        'Right = 3

        'Store his images
        SpriteImage(0) = New Bitmap("NPCs\" & ImageFolder & "\up.bmp")
        SpriteImage(1) = New Bitmap("NPCs\" & ImageFolder & "\down.bmp")
        SpriteImage(2) = New Bitmap("NPCs\" & ImageFolder & "\left.bmp")
        SpriteImage(3) = New Bitmap("NPCs\" & ImageFolder & "\right.bmp")

        'Make Lime his transparent color
        SpriteImage(0).MakeTransparent(Color.Lime)
        SpriteImage(1).MakeTransparent(Color.Lime)
        SpriteImage(2).MakeTransparent(Color.Lime)
        SpriteImage(3).MakeTransparent(Color.Lime)
    End Sub

    Public ReadOnly Property CurrentImage() As Bitmap
        Get
            Return SpriteImage(Direction)
        End Get
    End Property

    Public Sub UpdateMap()
        GameClass.map.Passable(Position.X, Position.Y) = False
    End Sub

End Class
Public Class clsEvent
    Dim i As Integer
    Dim Entry As DictionaryEntry
    Public Sub CheckForEvents()
        If Not GameClass.alex.InConversation Then
            For Each Entry In GameClass.NPCs
                Select Case GameClass.alex.Direction
                    Case clsSprite.Dir.Up
                        If GameClass.alex.TilePos.X = DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Position.X And GameClass.alex.TilePos.Y - 1 = DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Position.Y Then

                            DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Direction = clsSprite.Dir.Down
                            clsText.Text = "Hi, I'm the guy at " & DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Position.ToString()
                            GameClass.GameForm.Invalidate()
                            GameClass.alex.InConversation = True
                        End If
                    Case clsSprite.Dir.Down
                        If GameClass.alex.TilePos.X = DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Position.X And GameClass.alex.TilePos.Y + 1 = DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Position.Y Then

                            DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Direction = clsSprite.Dir.Up
                            clsText.Text = "Hi, I'm the guy at " & DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Position.ToString()
                            GameClass.GameForm.Invalidate()
                            GameClass.alex.InConversation = True
                        End If
                    Case clsSprite.Dir.Left
                        If GameClass.alex.TilePos.X - 1 = DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Position.X And GameClass.alex.TilePos.Y = DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Position.Y Then
                            DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Direction = clsSprite.Dir.Right
                            clsText.Text = "Hi, I'm the guy at " & DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Position.ToString()
                            GameClass.GameForm.Invalidate()
                            GameClass.alex.InConversation = True
                        End If
                    Case clsSprite.Dir.Right
                        If GameClass.alex.TilePos.X + 1 = DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Position.X And GameClass.alex.TilePos.Y = DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Position.Y Then
                            DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Direction = clsSprite.Dir.Left
                            clsText.Text = "Hi, I'm the guy at " & DirectCast(GameClass.NPCs(Entry.Key), clsNPC).Position.ToString()
                            GameClass.GameForm.Invalidate()
                            GameClass.alex.InConversation = True
                        End If
                End Select
            Next
        ElseIf GameClass.alex.InConversation Then
            'Since in this tutorial, we're not going to do much with text,
            'We'll just say that text is nothing
            clsText.Text = Nothing
            GameClass.GameForm.Invalidate()
            GameClass.alex.InConversation = False
            'In a future tutorial, we'll say something like
            ' "If there's more conversation remaining then advance the conversation "
            ' "Else, end conversation
        End If
    End Sub
End Class
Public Class clsText
    Public Shared Text As String
    Public Shared Position As Point = New Point(140, 395)
End Class
'A level file will look like:
'--------------------
'mapfile.map
'TileSet Name
'x,y <-- where x and y represent the character's tilepos
'NPCs:4   // The number of NPCs can be anything
'name of npc, folder of NPC images, x,y, direction, dialogue location
'name of npc, folder of NPC images, x,y, direction, dialogue location
'name of npc, folder of NPC images, x,y, direction, dialogue location
'name of npc, folder of NPC images, x,y, direction, dialogue location
Public Class clsLevel
    Dim Reader As System.IO.StreamReader
    Public Sub LoadLevel(ByVal Level As String)
        Reader = New System.IO.StreamReader("Levels/" & Level)
        Dim map, tileset As String
        'Read the map
        map = Reader.ReadLine
        'Read the tileset
        tileset = Reader.ReadLine
        'Now create the map
        GameClass.map = New clsMap(map, tileset)
        Dim s() As String
        'Now read the position (the line is written as "x,y")
        s = Reader.ReadLine.Split(",")
        'Since it is the tilepos, we multiply by 30
        GameClass.alex.Position.X = CInt(s(0)) * 30
        GameClass.alex.Position.Y = CInt(s(1)) * 30
        'Read the line that says "NPCs:4" (where 4 is any number)
        'and get that number
        s = Reader.ReadLine.Split(":")
        Dim NumberOfNPCs As Integer
        NumberOfNPCs = CInt(s(1))


        Dim x As Integer
        Dim dir As String
        Dim direction As clsSprite.Dir
        'Loop through and add each and every NPC
        For x = 0 To (NumberOfNPCs - 1)
            'Now split all the NPC data. For your reference:
            's(0) = name
            's(1) = folder
            's(2) = x position
            's(3) = y position
            's(4) = direction
            's(5) = dialogue location
            s = Reader.ReadLine.Split(",")
            'Get the direction (5th item when string is split by the comma)
            dir = s(4)
            'Now "translate" the string direction into a specific direction (clssprite.dir)

            Select Case dir
                Case "Down"
                    direction = clsSprite.Dir.Down
                Case "Up"
                    direction = clsSprite.Dir.Up
                Case "Left"
                    direction = clsSprite.Dir.Left
                Case "Right"
                    direction = clsSprite.Dir.Right
            End Select

            'Now create the NPC (this is a very LONG/CONFUSING line!)
            GameClass.NPCs.Add(s(0), New clsNPC(s(1), New Point(CInt(s(2)), CInt(s(3))), direction, s(5)))
        Next
        'The NPC will be rendered shortly after (since you added it to the NPC list), so invalidate
        GameClass.GameForm.Invalidate()
        'Close the file
        Reader.Close()
    End Sub
End Class