Imports System.CodeDom.Compiler
Imports System.Reflection

Public Class Form1

    Private Sub Button_Click(sender As Object, e As EventArgs) Handles Button0.Click, Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click, ButtonADDI.Click, ButtonBACK.Click, ButtonC.Click, ButtonDIVI.Click, ButtonMULT.Click, ButtonPLUSMOINS.Click, ButtonPOINT.Click, ButtonSOUS.Click, ButtonTOTAL.Click
        DisplayCaracter(sender)
    End Sub


    Private Sub DisplayCaracter(btn As Button)
        Select Case btn.Name
            Case "Button0"
                txtOperation.Text = txtOperation.Text & "0"
            Case "Button1"
                txtOperation.Text = txtOperation.Text & "1"
            Case "Button2"
                txtOperation.Text = txtOperation.Text & "2"
            Case "Button3"
                txtOperation.Text = txtOperation.Text & "3"
            Case "Button4"
                txtOperation.Text = txtOperation.Text & "4"
            Case "Button5"
                txtOperation.Text = txtOperation.Text & "5"
            Case "Button6"
                txtOperation.Text = txtOperation.Text & "6"
            Case "Button7"
                txtOperation.Text = txtOperation.Text & "7"
            Case "Button8"
                txtOperation.Text = txtOperation.Text & "8"
            Case "Button9"
                txtOperation.Text = txtOperation.Text & "9"
            Case "ButtonADDI"
                txtOperation.Text = txtOperation.Text & "+"
            Case "ButtonSOUS"
                txtOperation.Text = txtOperation.Text & "-"
            Case "ButtonDIVI"
                txtOperation.Text = txtOperation.Text & "/"
            Case "ButtonMULT"
                txtOperation.Text = txtOperation.Text & "*"
            Case "ButtonPOINT"
                txtOperation.Text = txtOperation.Text & "."
            Case "ButtonPLUSMOINS"
                txtOperation.Text = "-" & txtOperation.Text
            Case "ButtonC"
                txtOperation.Text = ""
                txtTOTAL.Text = ""
                txtOperation.Focus()
            Case "ButtonBACK"
                If Len(txtOperation.Text) > 0 Then
                    txtOperation.Text = Mid(txtOperation.Text, 1, Len(txtOperation.Text) - 1)
                    txtOperation.SelectionStart = Len(txtOperation.Text)
                    txtOperation.Focus()
                End If
            Case "ButtonTOTAL"
                txtTOTAL.Text = TransformOperation(Trim(Replace(txtOperation.Text, " ", "")))
                txtOperation.Focus()
                txtOperation.SelectionStart = Len(txtOperation.Text)
                '   txtOperation.Text = ""
        End Select

    End Sub

    Private Sub txtOperation_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtOperation.KeyPress
        'empeche la clé enter de changer de ligne 
        If e.KeyChar = Chr(13) Then
            e.KeyChar = ChrW(Keys.None)
        End If
    End Sub

    Private Sub txtOperation_KeyUp(sender As Object, e As KeyEventArgs) Handles txtOperation.KeyUp
        'récupère les clés du clavier
        Select Case e.KeyCode
            Case Keys.Enter
                Button_Click(ButtonTOTAL, EventArgs.Empty)
            Case Keys.Delete
                Button_Click(ButtonC, EventArgs.Empty)
            Case Else
                If Len(txtOperation.Text) = 35 Then MsgBox("35 caracters maximum")
                If Len(txtOperation.Text) = 0 Then txtTOTAL.Text = ""
        End Select

    End Sub

    Private Function TransformOperation(ByVal txt As String) As String
        Dim vbCProvider As New VBCodeProvider
        Dim comp As New CompilerParameters 'nouvel objet compilateur
        comp.GenerateExecutable = False
        comp.GenerateInMemory = True

        Dim TempModuleSource As String = "Imports System" & Environment.NewLine &
             "Namespace ns " & Environment.NewLine &
            "Public Class clComp" & Environment.NewLine &
            "Public Shared Function Evaluate()" & Environment.NewLine &
            "Return " & txt & Environment.NewLine &
            "End Function" & Environment.NewLine &
            "End Class" & Environment.NewLine &
            "End Namespace"

        'invoque la class compilateur et appelle sa fonction evaluate ou récupére son erreur
        Dim compRes As CompilerResults = vbCProvider.CompileAssemblyFromSource(comp, TempModuleSource)
        If compRes.Errors.Count > 0 Then
            Return "Erreur*"
        Else
            Dim mInfo As MethodInfo = compRes.CompiledAssembly.GetType("ns.clComp").GetMethod("Evaluate")
            Return Convert.ToString(mInfo.Invoke(Nothing, Nothing))
        End If
    End Function


End Class
