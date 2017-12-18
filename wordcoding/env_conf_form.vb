Public Class env_conf_form
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        env_conf.win_work_path = Me.TextBox1.Text
        env_conf.linux_ip = Me.TextBox2.Text
        env_conf.linux_user = Me.TextBox3.Text
        env_conf.linux_password = Me.TextBox4.Text
        env_conf.linux_tran = Me.TextBox5.Text
        env_conf.linux_dir = Me.TextBox6.Text
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        env_conf.win_work_path = Me.TextBox1.Text
        env_conf.linux_ip = Me.TextBox2.Text
        env_conf.linux_user = Me.TextBox3.Text
        env_conf.linux_password = Me.TextBox4.Text
        env_conf.linux_tran = Me.TextBox5.Text
        env_conf.linux_dir = Me.TextBox6.Text
        Me.Close()
    End Sub

End Class