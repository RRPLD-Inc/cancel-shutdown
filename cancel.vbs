Set objShell = CreateObject("WScript.Shell")



Sub RunCommand()

    ' Run the "shutdown /a" command to abort shutdown

    objShell.Run "shutdown /a", 0, False

End Sub



Sub RepeatCommand()

    ' Run the command initially

    RunCommand

    

    ' Repeat the command every 1s

    Do

        ' Sleep for 1s (1,000 milliseconds)

        WScript.Sleep 1000

        

        ' Run the command again

        RunCommand

    Loop

End Sub



' Start repeating the command

RepeatCommand