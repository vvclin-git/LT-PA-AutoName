Imports LightTools
Imports LTLOCATORLib
Class MainWindow
    Dim lt As LTAPI
    Dim js As New LTCOM64.JSNET
    Dim pID As Integer
    Dim stat As LTReturnCodeEnum
    Public Sub New()

        ' 設計工具需要此呼叫。
        InitializeComponent()

        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。
        lt = LT_Getter.getLTAPIServer
        If IsNothing(lt) Then
            MsgBox("LightTools session not found!")
        Else
            Me.Title = "Pickups and Alias Auto Naming (PID = " + Str(lt.GetServerID()) + ")"
        End If


    End Sub

    Private Sub autoName()
        Dim aliasListKey, pickupListKey As String
        Dim aliasKey, pickupKey As String
        Dim newName As String
        Dim stat As LTReturnCodeEnum
        'get the keys of alias and pickups lists
        aliasListKey = lt.DbList("LENS_MANAGER[1].PARAMETRIC_CONTROLS[Parametric_Controls]", "ALIAS")
        pickupListKey = lt.DbList("LENS_MANAGER[1].PARAMETRIC_CONTROLS[Parametric_Controls]", "PICKUP")
        'loop through the list and rename
        Do
            aliasKey = lt.ListNext(aliasListKey, stat)
            'check if the end of the list is reached
            If stat = 53 Then
                Exit Do
            End If
            If Microsoft.VisualBasic.Left(lt.DbGet(aliasKey, "Name"), 1) <> "_" Then
                newName = getNewName(lt.DbGet(aliasKey, "ValueStrUI"))
                lt.DbSet(aliasKey, "Name", newName)
            End If
        Loop
        Do
            pickupKey = lt.ListNext(pickupListKey, stat)
            'check if the end of the list is reached
            If stat = 53 Then
                Exit Do
            End If
            If Microsoft.VisualBasic.Left(lt.DbGet(pickupKey, "Name"), 1) <> "_" Then
                newName = getNewName(lt.DbGet(pickupKey, "ConnectionStringUI"))
                lt.DbSet(pickupKey, "Name", newName)
            End If
        Loop

    End Sub

    Private Function getName(accessStr As String) As String
        Return Replace(Split(Split(accessStr, "[")(1), "]")(0), " ", "_")
    End Function

    Private Function getTypeName(accessStr As String) As String
        Return Split(accessStr, "[")(0)
    End Function

    Private Function getNewName(valueStr As String) As String
        'rename the pickups/alias according to a specific naming scheme
        Dim valueStrArray() As String
        Dim entName, newName As String
        valueStr = Replace(valueStr, vbCrLf, "")
        valueStrArray = Split(valueStr, ".")
        entName = getName(valueStrArray(UBound(valueStrArray) - 1)) 'stuff outside of the name
        Select Case valueStrArray(1)
            Case "PROPERTY_MANAGER[Optical Properties Manager]"
                entName = getName(valueStrArray(2)) + "__" + getName(valueStrArray(3))
            Case "USER_MATERIALS[User Materials]"
                entName = getName(valueStrArray(2)) + "__" + getName(valueStrArray(3))
            Case "ILLUM_MANAGER[Illumination Manager]"
                If Trim(getTypeName(valueStrArray(3))) Like "*SURFACE_RECEIVER*" Then
                    entName = getName(valueStrArray(3)) + "__" + getName(valueStrArray(4)) + "__" + getName(valueStrArray(5))
                End If
        End Select
        newName = "_" + entName + "__" + valueStrArray(UBound(valueStrArray))
        getNewName = newName
    End Function

    Private Sub autoNameBtn_Click(sender As Object, e As RoutedEventArgs) Handles autoNameBtn.Click
        Call autoName()
    End Sub
End Class
