Module send_file_to_linux
    Sub send_file_to_linux()
        Dim docpath As String
        Dim cmd As String
        'docpath = env_conf.win_work_path & "\send_file_to_yc"
        docpath = env_conf.win_sys_dir
        Call ChDrive(docpath)
        Call ChDir(docpath)
        cmd = docpath & "\file_bat\mmput.bat"
        Call Shell(cmd)
    End Sub

End Module
