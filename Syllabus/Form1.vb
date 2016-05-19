Public Class Form1

    Dim Game As GameClass = New GameClass(Me)
    Dim x, y As Integer
    Dim SetRes As ScreenSettings


  


    Private Sub Form1_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Game.Terminate()



    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim result As MsgBoxResult
        result = MessageBox.Show("Do you really want to quit?", "Quit?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = MsgBoxResult.No Then
            e.Cancel = True
        End If

    End Sub




    Private Sub Form1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
       

    End Sub

    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load


        Game.Level.LoadLevel("level1.lvl")



    End Sub

  
   

    Private Sub Form1_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint


        Me.SetStyle(ControlStyles.DoubleBuffer, True)
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        Me.SetStyle(ControlStyles.UserPaint, True)
        Game.Paint(sender, e)




    End Sub




End Class
