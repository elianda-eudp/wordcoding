Module get_file_from_linux
    Sub get_file_from_linux()
        Dim docpath As String
        Dim cmd As String

        docpath = env_conf.win_sys_dir
        cmd = docpath & "\file_bat\mmget.bat"
        Call ChDrive(docpath)
        Call ChDir(docpath)
        'a_path = System.Windows.Forms.Application.StartupPath()
        'MsgBox(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase)
        'MsgBox(System.Windows.Forms.Application.ExecutablePath)
        'MsgBox(System.Reflection.Assembly.GetExecutingAssembly().Location)
        'MsgBox(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
        'MsgBox(System.AppDomain.CurrentDomain.BaseDirectory)
        'MsgBox(docpath)
        'MsgBox(a_path)
        'MsgBox(Environ("path"))
        'MsgBox(Environ("ProductName"))

        'cmd = "winscp.exe    /console /command " + """Option batch Continue""" + " /script=.\getpsftp.dat"
        MsgBox(cmd)
        Call Shell(cmd)
        'Call Shell("winscp.exe    /console /command "Option batch Continue" /script=.\getpsftp.dat")

    End Sub

End Module
