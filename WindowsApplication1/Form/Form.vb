Public Class Form
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Timer1.Start()
        Else
            Timer1.Stop()
        End If
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim prs As Process() = Process.GetProcessesByName("")
        'Vai checar se o processo esta aberto'
        If prs.Length > 0 Then 'sim esta aberto'
            On Error Resume Next
        Else
            MsgBox("Processo está fechado")
            'Close()
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        WriteDMAFloat("hl.exe", &H25069BC, Offsets:={&H88}, Value:=-2125.098145, Level:=1, nsize:=4) ' cordenada Y
        WriteDMAFloat("hl.exe", &H25069BC, Offsets:={&H8C}, Value:=971.6911621, Level:=1, nsize:=4) ' cordenada X
        WriteDMAFloat("hl.exe", &H25069BC, Offsets:={&H90}, Value:=-187.96875, Level:=1, nsize:=4) ' cordenada Z
        'Iniciando Trainer
    End Sub
End Class
