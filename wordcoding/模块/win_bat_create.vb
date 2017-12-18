Module win_bat_create
    Sub win_bat_create()


        Dim docpath As Object
        'docpath = System.Windows.Forms.Application.StartupPath() & "\"
        docpath = env_conf.win_sys_dir & "\"
        If Dir(docpath & "file_bat", vbDirectory) = "" Then
            MkDir(docpath & "file_bat")
        End If

        FileClose(15)
        FileOpen(15, docpath & "file_bat\mmget.bat", OpenMode.Output)

        Print(15, "cd " & docpath & "file_bat\" & vbCrLf)
        Print(15, "" & vbCrLf)
        Print(15, "winscp.exe    /console /command " + """Option batch Continue""" + " /script=.\getpsftp.dat" & vbCrLf)
        Print(15, "" & vbCrLf)
        FileClose(15)

        'FileOpen(15, docpath & "file_bat\getpsftp.dat", OpenMode.Output)

        'Print(15, "option echo on" & vbCrLf)
        'Print(15, "option transfer ascii" & vbCrLf)
        'Print(15, "open " & env_conf.linux_user & ":" & env_conf.linux_password & "@" & env_conf.linux_ip & vbCrLf)

        'Print(15, "cd /home/" & env_conf.linux_user & "/vba_tool/src" & vbCrLf)
        'Print(15, "lcd " & env_conf.win_work_path & "\src\remote_host" & vbCrLf)
        'Print(15, "get *_handle.c" & vbCrLf)
        'Print(15, "exit" & vbCrLf)
        'Print(15, "" & vbCrLf)
        'FileClose(15)
        '生成utf-8

        Dim gbk_text As String
        gbk_text = "option echo on" & vbCrLf
        gbk_text = gbk_text & "option transfer ascii" & vbCrLf
        gbk_text = gbk_text & "open " & env_conf.linux_user & ":" & env_conf.linux_password & "@" & env_conf.linux_ip & vbCrLf
        gbk_text = gbk_text & "cd /home/" & env_conf.linux_user & "/vba_tool/src" & vbCrLf
        gbk_text = gbk_text & "lcd " & env_conf.win_work_path & "\src\remote_host" & vbCrLf
        gbk_text = gbk_text & "get *_handle.c" & vbCrLf
        gbk_text = gbk_text & "exit" & vbCrLf
        gbk_text = gbk_text & "" & vbCrLf
        Call SaveUTF8(gbk_text, docpath & "file_bat\getpsftp.dat")


        'FileOpen(15, docpath & "file_bat\psftp.bat", OpenMode.Output)
        'Print(15, "option echo on" & vbCrLf)
        'Print(15, "option transfer ascii" & vbCrLf)
        'Print(15, "open " & env_conf.linux_user & ":" & env_conf.linux_password & "@" & env_conf.linux_ip & vbCrLf)

        'Print(15, "cd /home/" & env_conf.linux_user & "/db_ctl/progcfg" & vbCrLf)
        'Print(15, "put " & env_conf.win_work_path & "\dbcfg\*  -neweronly" & vbCrLf)
        'Print(15, "" & vbCrLf)

        'Print(15, "cd /home/" & env_conf.linux_user & "/db_ctl/inc" & vbCrLf)
        'Print(15, "put " & env_conf.win_work_path & "\src\*.h -neweronly" & vbCrLf)
        'Print(15, "" & vbCrLf)

        'Print(15, "cd /home/" & env_conf.linux_user & "/db_ctl/src" & vbCrLf)
        'Print(15, "put " & env_conf.win_work_path & "\src\*.c  -neweronly" & vbCrLf)
        'Print(15, "put " & env_conf.win_work_path & "\regedit\*_regedit.cfg  -neweronly" & vbCrLf)
        'Print(15, "" & vbCrLf)
        'Print(15, "exit" & vbCrLf)
        'FileClose(15)

        gbk_text = ""
        gbk_text = gbk_text & "option echo on" & vbCrLf
        gbk_text = gbk_text & "option transfer ascii" & vbCrLf
        gbk_text = gbk_text & "open " & env_conf.linux_user & ":" & env_conf.linux_password & "@" & env_conf.linux_ip & vbCrLf
        gbk_text = gbk_text & "cd /home/" & env_conf.linux_user & "/db_ctl/progcfg" & vbCrLf
        gbk_text = gbk_text & "put " & env_conf.win_work_path & "\dbcfg\*  -neweronly" & vbCrLf
        gbk_text = gbk_text & "" & vbCrLf
        gbk_text = gbk_text & "cd /home/" & env_conf.linux_user & "/db_ctl/inc" & vbCrLf
        gbk_text = gbk_text & "put " & env_conf.win_work_path & "\src\*.h -neweronly" & vbCrLf
        gbk_text = gbk_text & "" & vbCrLf
        gbk_text = gbk_text & "cd /home/" & env_conf.linux_user & "/db_ctl/src" & vbCrLf
        gbk_text = gbk_text & "put " & env_conf.win_work_path & "\src\*.c  -neweronly" & vbCrLf
        gbk_text = gbk_text & "put " & env_conf.win_work_path & "\regedit\*_regedit.cfg  -neweronly" & vbCrLf
        gbk_text = gbk_text & "" & vbCrLf
        gbk_text = gbk_text & "exit" & vbCrLf
        Call SaveUTF8(gbk_text, docpath & "file_bat\psftp.bat")

        'FileOpen(15, docpath & "file_bat\crt.bat", OpenMode.Output)

        'Print(15, "#!/bin/bash –login" & vbCrLf)
        'Print(15, ". .bash_profile" & vbCrLf)
        'Print(15, "cd /home/financialsys/db_ctl/progcfg" & vbCrLf)
        'Print(15, "for LINE in `cat /home/financialsys/vba_tool/src/prog_regedit.cfg`" & vbCrLf)
        'Print(15, "do" & vbCrLf)
        'Print(15, "echo $LINE" & vbCrLf)
        'Print(15, "prog_dbcfg_to_table_cfg -p $LINE" & vbCrLf)
        'Print(15, "done" & vbCrLf)
        'Print(15, "" & vbCrLf)

        'Print(15, "cd /home/financialsys/db_ctl/pccfg" & vbCrLf)
        'Print(15, "for  LINE in `cat /home/financialsys/vba_tool/src/tables_regedit.cfg`" & vbCrLf)
        'Print(15, "do" & vbCrLf)
        'Print(15, "echo $LINE" & vbCrLf)
        'Print(15, "crt_pcfiles  $LINE" & vbCrLf)
        'Print(15, "done" & vbCrLf)
        'Print(15, "" & vbCrLf)

        'Print(15, "cd /home/financialsys/vba_tool/libsrc" & vbCrLf)
        'Print(15, "touch -d " & """5 days ago""" & " *.c" & vbCrLf)
        'Print(15, "pwd" & vbCrLf)
        'Print(15, "make" & vbCrLf)
        'Print(15, "" & vbCrLf)

        'Print(15, "cd /home/financialsys/vba_tool/src" & vbCrLf)
        'Print(15, "touch -d " & """5 days ago""" & " *.c" & vbCrLf)
        'Print(15, "pwd" & vbCrLf)
        'Print(15, "make" & vbCrLf)
        'Print(15, "" & vbCrLf)
        'FileClose(15)

        gbk_text = ""
        gbk_text = gbk_text & "option echo on" & vbCrLf
        gbk_text = gbk_text & "#!/bin/bash –login" & vbCrLf
        gbk_text = gbk_text & ". .bash_profile" & vbCrLf
        gbk_text = gbk_text & "cd /home/financialsys/db_ctl/progcfg" & vbCrLf
        gbk_text = gbk_text & "for LINE in `cat /home/financialsys/vba_tool/src/prog_regedit.cfg`" & vbCrLf
        gbk_text = gbk_text & "do" & vbCrLf
        gbk_text = gbk_text & "echo $LINE" & vbCrLf
        gbk_text = gbk_text & "prog_dbcfg_to_table_cfg -p $LINE" & vbCrLf
        gbk_text = gbk_text & "done" & vbCrLf
        gbk_text = gbk_text & "" & vbCrLf

        gbk_text = gbk_text & "cd /home/financialsys/db_ctl/pccfg" & vbCrLf
        gbk_text = gbk_text & "for  LINE in `cat /home/financialsys/vba_tool/src/tables_regedit.cfg`" & vbCrLf
        gbk_text = gbk_text & "do" & vbCrLf
        gbk_text = gbk_text & "echo $LINE" & vbCrLf
        gbk_text = gbk_text & "crt_pcfiles  $LINE" & vbCrLf
        gbk_text = gbk_text & "done" & vbCrLf
        gbk_text = gbk_text & "" & vbCrLf

        gbk_text = gbk_text & "cd /home/financialsys/vba_tool/libsrc" & vbCrLf
        gbk_text = gbk_text & "touch -d " & """5 days ago""" & " *.c" & vbCrLf
        gbk_text = gbk_text & "pwd" & vbCrLf
        gbk_text = gbk_text & "make" & vbCrLf
        gbk_text = gbk_text & "" & vbCrLf

        gbk_text = gbk_text & "cd /home/financialsys/vba_tool/src" & vbCrLf
        gbk_text = gbk_text & "touch -d " & """5 days ago""" & " *.c" & vbCrLf
        gbk_text = gbk_text & "pwd" & vbCrLf
        gbk_text = gbk_text & "make" & vbCrLf
        gbk_text = gbk_text & "" & vbCrLf
        Call SaveUTF8(gbk_text, docpath & "file_bat\crt.bat")

        FileOpen(15, docpath & "file_bat\mmput.bat", OpenMode.Output)

        'Print(15, "cd C:\Program Files (x86)\WinSCP" & vbCrLf)
        Print(15, "cd " & docpath & "file_bat\" & vbCrLf)
        Print(15, "" & vbCrLf)
        Print(15, "winscp.exe    /console /command " + """Option batch Continue""" + " /script=.\psftp.dat" & vbCrLf)
        Print(15, "plink.exe -ssh -v -pw " & env_conf.linux_user & " " & env_conf.linux_password & "@" & env_conf.linux_ip & " -m .\crt.bat" & vbCrLf)
        'Print(15, "putty.exe -l " & env_conf.linux_user & " -pw " & env_conf.linux_password & " " & env_conf.linux_ip & vbCrLf)
        Print(15, "" & vbCrLf)
        FileClose(15)
    End Sub
End Module
