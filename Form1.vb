Imports System.IO
Imports System.Security.Cryptography
Public Class Form1
    Dim AppName As String = "Team Builder by Cat"
    Dim DBFilename As String = "Pokedata.sdf"
    Dim Connection As New System.Data.SqlServerCe.SqlCeConnection("Data Source=" & DBFilename)
    Dim pkm1_selected As UShort = 0
    Dim pkm2_selected As UShort = 0
    Dim pkm3_selected As UShort = 0
    Dim pkm4_selected As UShort = 0
    Dim pkm5_selected As UShort = 0
    Dim pkm6_selected As UShort = 0
    Dim pokemonNames(5) As String
    Dim filename As String = ""
    Dim prettyFilename As String = ""

    Dim pkm1_nature_multipliers() As Double = {1.0, 1.0, 1.0, 1.0, 1.0}
    Dim pkm2_nature_multipliers() As Double = {1.0, 1.0, 1.0, 1.0, 1.0}
    Dim pkm3_nature_multipliers() As Double = {1.0, 1.0, 1.0, 1.0, 1.0}
    Dim pkm4_nature_multipliers() As Double = {1.0, 1.0, 1.0, 1.0, 1.0}
    Dim pkm5_nature_multipliers() As Double = {1.0, 1.0, 1.0, 1.0, 1.0}
    Dim pkm6_nature_multipliers() As Double = {1.0, 1.0, 1.0, 1.0, 1.0}


    Public Enum Stat As Byte
        HP = 0 : Attack = 1 : Defense = 2
        SpecialAttack = 3 : SpecialDefense = 4 : Speed = 5
    End Enum

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox(AppName & ", v" & Me.ProductVersion.Substring(0, Me.ProductVersion.IndexOf(".", Me.ProductVersion.IndexOf(".", 0) + 1)) & vbCrLf & vbCrLf & _
               "Created by Cat333Pokémon for Victory Road." & vbCrLf & vbCrLf & _
               "This product uses data by the free veekun project.", _
                MsgBoxStyle.Information)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim suffix As String = ""



        If Not FileIO.FileSystem.FileExists(DBFilename) Then
            MsgBox("The database '" & DBFilename & "' cannot be found. Please try re-downloading the program, extracting the files, or locating the database. Be sure '" & DBFilename & "' is in the same folder as this program.", MsgBoxStyle.Critical)
            End
        Else
            'Dim ourHash(0) As Byte
            Try
                Dim attribute As System.IO.FileAttributes = FileAttributes.Normal
                File.SetAttributes(DBFilename, attribute)

                '    Dim ourHashAlg As HashAlgorithm = HashAlgorithm.Create("MD5")
                '    Dim fileToHash As FileStream = New FileStream("Pokedata.sdf", FileMode.Open, FileAccess.Read)

                '    ourHash = ourHashAlg.ComputeHash(fileToHash)
                '    fileToHash.Close()
                '    If Convert.ToBase64String(ourHash) <> "lfDIoAUEhDD/aavQLzKICw==" Then
                '        MsgBox("The database '" & DBFilename & "' is corrupt. Please try re-downloading it: " & vbCrLf & vbCrLf & Convert.ToBase64String(ourHash), MsgBoxStyle.Critical)
                '        End
                '    End If
            Catch ex As IOException
                MsgBox("The database '" & DBFilename & "' cannot be opened: " & ex.Message, MsgBoxStyle.Critical)
                End
            End Try
        End If


        'Add hash check here to make sure the RIGHT file is loaded

        Try
            Connection.Open()
        Catch ex As Exception
            MsgBox("Could not connect to the database '" & DBFilename & "':" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical)
            End
        End Try

        'Get Pokémon list
        Dim Command As New System.Data.SqlServerCe.SqlCeCommand("SELECT species_id, name, form_identifier FROM pokemon", Connection)
        Dim myReader As System.Data.SqlServerCe.SqlCeDataReader = Command.ExecuteReader(CommandBehavior.Default)

        While myReader.Read()
            suffix = ""
            If Not myReader.IsDBNull(2) Then
                suffix = " (" & myReader.GetString(2) & ")"
            End If
            pkm1.Items.Add(myReader.GetString(1) & suffix)
            pkm2.Items.Add(myReader.GetString(1) & suffix)
            pkm3.Items.Add(myReader.GetString(1) & suffix)
            pkm4.Items.Add(myReader.GetString(1) & suffix)
            pkm5.Items.Add(myReader.GetString(1) & suffix)
            pkm6.Items.Add(myReader.GetString(1) & suffix)
        End While
        myReader.Close()

        'Get nature list
        Command = New System.Data.SqlServerCe.SqlCeCommand("SELECT name FROM natures", Connection)
        myReader = Command.ExecuteReader(CommandBehavior.Default)
        While myReader.Read()
            pkm1_nature.Items.Add(myReader.GetString(0))
            pkm2_nature.Items.Add(myReader.GetString(0))
            pkm3_nature.Items.Add(myReader.GetString(0))
            pkm4_nature.Items.Add(myReader.GetString(0))
            pkm5_nature.Items.Add(myReader.GetString(0))
            pkm6_nature.Items.Add(myReader.GetString(0))
        End While


        'Get move list
        Command = New System.Data.SqlServerCe.SqlCeCommand("SELECT name FROM moves", Connection)
        myReader = Command.ExecuteReader(CommandBehavior.Default)
        While myReader.Read()
            pkm1_move1.Items.Add(myReader.GetString(0))
            pkm1_move2.Items.Add(myReader.GetString(0))
            pkm1_move3.Items.Add(myReader.GetString(0))
            pkm1_move4.Items.Add(myReader.GetString(0))
            pkm2_move1.Items.Add(myReader.GetString(0))
            pkm2_move2.Items.Add(myReader.GetString(0))
            pkm2_move3.Items.Add(myReader.GetString(0))
            pkm2_move4.Items.Add(myReader.GetString(0))
            pkm3_move1.Items.Add(myReader.GetString(0))
            pkm3_move2.Items.Add(myReader.GetString(0))
            pkm3_move3.Items.Add(myReader.GetString(0))
            pkm3_move4.Items.Add(myReader.GetString(0))
            pkm4_move1.Items.Add(myReader.GetString(0))
            pkm4_move2.Items.Add(myReader.GetString(0))
            pkm4_move3.Items.Add(myReader.GetString(0))
            pkm4_move4.Items.Add(myReader.GetString(0))
            pkm5_move1.Items.Add(myReader.GetString(0))
            pkm5_move2.Items.Add(myReader.GetString(0))
            pkm5_move3.Items.Add(myReader.GetString(0))
            pkm5_move4.Items.Add(myReader.GetString(0))
            pkm6_move1.Items.Add(myReader.GetString(0))
            pkm6_move2.Items.Add(myReader.GetString(0))
            pkm6_move3.Items.Add(myReader.GetString(0))
            pkm6_move4.Items.Add(myReader.GetString(0))
        End While


        'Get item list
        Command = New System.Data.SqlServerCe.SqlCeCommand("SELECT name FROM items", Connection)
        myReader = Command.ExecuteReader(CommandBehavior.Default)
        While myReader.Read()
            pkm1_helditem.Items.Add(myReader.GetString(0))
            pkm2_helditem.Items.Add(myReader.GetString(0))
            pkm3_helditem.Items.Add(myReader.GetString(0))
            pkm4_helditem.Items.Add(myReader.GetString(0))
            pkm5_helditem.Items.Add(myReader.GetString(0))
            pkm6_helditem.Items.Add(myReader.GetString(0))
        End While


        Connection.Close()

        pkm1.SelectedIndex = 0
        pkm1_nature.SelectedIndex = 0
        pkm1_move1.SelectedIndex = 0
        pkm1_move2.SelectedIndex = 0
        pkm1_move3.SelectedIndex = 0
        pkm1_move4.SelectedIndex = 0
        pkm1_helditem.SelectedIndex = 0
        pkm1_gender.SelectedIndex = 0

        pkm2.SelectedIndex = 0
        pkm2_nature.SelectedIndex = 0
        pkm2_move1.SelectedIndex = 0
        pkm2_move2.SelectedIndex = 0
        pkm2_move3.SelectedIndex = 0
        pkm2_move4.SelectedIndex = 0
        pkm2_helditem.SelectedIndex = 0
        pkm2_gender.SelectedIndex = 0

        pkm3.SelectedIndex = 0
        pkm3_nature.SelectedIndex = 0
        pkm3_move1.SelectedIndex = 0
        pkm3_move2.SelectedIndex = 0
        pkm3_move3.SelectedIndex = 0
        pkm3_move4.SelectedIndex = 0
        pkm3_helditem.SelectedIndex = 0
        pkm3_gender.SelectedIndex = 0

        pkm4.SelectedIndex = 0
        pkm4_nature.SelectedIndex = 0
        pkm4_move1.SelectedIndex = 0
        pkm4_move2.SelectedIndex = 0
        pkm4_move3.SelectedIndex = 0
        pkm4_move4.SelectedIndex = 0
        pkm4_helditem.SelectedIndex = 0
        pkm4_gender.SelectedIndex = 0

        pkm5.SelectedIndex = 0
        pkm5_nature.SelectedIndex = 0
        pkm5_move1.SelectedIndex = 0
        pkm5_move2.SelectedIndex = 0
        pkm5_move3.SelectedIndex = 0
        pkm5_move4.SelectedIndex = 0
        pkm5_helditem.SelectedIndex = 0
        pkm5_gender.SelectedIndex = 0

        pkm6.SelectedIndex = 0
        pkm6_nature.SelectedIndex = 0
        pkm6_move1.SelectedIndex = 0
        pkm6_move2.SelectedIndex = 0
        pkm6_move3.SelectedIndex = 0
        pkm6_move4.SelectedIndex = 0
        pkm6_helditem.SelectedIndex = 0
        pkm6_gender.SelectedIndex = 0
    End Sub

    Function UpdateMon(ByRef SelectedIndex As UShort)
        Dim iteration As Byte = 0
        Dim suffix As String = ""
        Dim suffix2 As String = ""
        Dim pokedata(24) As Object
        Dim abilities(3) As String
        Dim pokemon_id As UShort = 0
        Dim efficacy1() As Byte = {100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100}
        Dim efficacy2() As Byte = {100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100}
        Dim efficacyNormalized() As UShort = {100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100}

        pokedata(0) = 0

        Try
            Connection.Open()
        Catch ex As Exception
            MsgBox("Could not connect to the database '" & DBFilename & "'.", MsgBoxStyle.Critical)
            Return pokedata
        End Try

        Dim Command As New System.Data.SqlServerCe.SqlCeCommand("SELECT species_id, file_suffix, genus, height, weight, hp, attack, defense, spatk, spdef, speed, entry_id, name, form_identifier, type1, type2 FROM pokemon WHERE pokemon.id = " & SelectedIndex, Connection)
        Dim myReader As System.Data.SqlServerCe.SqlCeDataReader = Command.ExecuteReader(CommandBehavior.Default)

        While myReader.Read()
            'Get species ID (typically 1-649) (0)
            pokedata(0) = myReader.GetInt16(0)

            'Get forme suffix
            If Not myReader.IsDBNull(1) Then
                suffix = "_" & myReader.GetString(1)
            End If

            'Get genus (1)
            pokedata(1) = myReader.GetString(2) & " Pokémon"

            'Get height (22) and weight (23)
            pokedata(22) = FormatNumber(CDec(myReader.GetInt16(3)) / 10, 1, TriState.True, TriState.UseDefault, TriState.True) & " m"
            pokedata(23) = FormatNumber(CDec(myReader.GetInt16(4)) / 10, 1, TriState.True, TriState.UseDefault, TriState.True) & " kg"

            'Get Pokémon stats (7-17, odd) and label widths (8-18, even)
            pokedata(7) = myReader.GetByte(5)
            pokedata(8) = 30 + Int((myReader.GetByte(5) / 255.0) * 100)
            pokedata(9) = myReader.GetByte(6)
            pokedata(10) = 30 + Int((myReader.GetByte(6) / 255.0) * 100)
            pokedata(11) = myReader.GetByte(7)
            pokedata(12) = 30 + Int((myReader.GetByte(7) / 255.0) * 100)
            pokedata(13) = myReader.GetByte(8)
            pokedata(14) = 30 + Int((myReader.GetByte(8) / 255.0) * 100)
            pokedata(15) = myReader.GetByte(9)
            pokedata(16) = 30 + Int((myReader.GetByte(9) / 255.0) * 100)
            pokedata(17) = myReader.GetByte(10)
            pokedata(18) = 30 + Int((myReader.GetByte(10) / 255.0) * 100)

            'Sum of stats (19)
            pokedata(19) = Int(myReader.GetByte(5)) + Int(myReader.GetByte(6)) + _
                Int(myReader.GetByte(7)) + Int(myReader.GetByte(8)) + _
                Int(myReader.GetByte(9)) + Int(myReader.GetByte(10))

            pokemon_id = myReader.GetInt16(11)

            'Name with suffix (21)
            suffix2 = ""
            If Not myReader.IsDBNull(13) Then
                suffix2 = " (" & myReader.GetString(13) & ")"
            End If
            pokedata(21) = myReader.GetInt16(0) & ". " & myReader.GetString(12) & suffix2

            'Name without suffix (5)
            pokedata(5) = myReader.GetString(12)

            'Get types (2, 3)
            pokedata(2) = myReader.GetByte(14)
            If Not myReader.IsDBNull(15) Then
                pokedata(3) = myReader.GetByte(15)
            Else
                pokedata(3) = 0
            End If
        End While
        myReader.Close()

        'Get resource filename for image (4)
        pokedata(4) = "_" & pokedata(0) & suffix

        ''Get type 1 (2,5)
        'Command = New System.Data.SqlServerCe.SqlCeCommand("SELECT types.type_id, type_name FROM types, pokemon WHERE types.type_id = pokemon.type1 AND pokemon.id = " & SelectedIndex, Connection)
        'myReader = Command.ExecuteReader(CommandBehavior.Default)

        'While myReader.Read()
        '    pokedata(2) = myReader.GetByte(0)
        '    pokedata(5) = myReader.GetString(1)
        'End While
        'myReader.Close()

        ''Get type 2 (3,6)
        'Command = New System.Data.SqlServerCe.SqlCeCommand("SELECT types.type_id, type_name FROM types, pokemon WHERE types.type_id = pokemon.type2 AND pokemon.id = " & SelectedIndex, Connection)
        'myReader = Command.ExecuteReader(CommandBehavior.Default)

        'While myReader.Read()
        '    pokedata(3) = myReader.GetByte(0)
        '    pokedata(6) = myReader.GetString(1)
        'End While
        'myReader.Close()

        'Get abilities (20)
        Command = New System.Data.SqlServerCe.SqlCeCommand("SELECT ability_names.name FROM ability_names, pokemon_abilities WHERE ability_names.ability_id = pokemon_abilities.ability_id AND pokemon_abilities.pokemon_id = " & pokemon_id & " ORDER BY slot ASC", Connection)
        myReader = Command.ExecuteReader(CommandBehavior.Default)

        iteration = 0
        While myReader.Read()
            abilities(iteration) = myReader.GetString(0)
            iteration += 1
        End While
        myReader.Close()

        pokedata(20) = abilities

        'Get type effectiveness (24)
        If pokedata(2) <> 0 Then
            Command = New System.Data.SqlServerCe.SqlCeCommand("SELECT damage_factor FROM type_efficacy WHERE target_type_id = " & pokedata(2) & " ORDER BY damage_type_id ASC", Connection)
            myReader = Command.ExecuteReader(CommandBehavior.Default)
            iteration = 0
            While myReader.Read()
                efficacy1(iteration) = myReader.GetByte(0)
                iteration += 1
            End While
            myReader.Close()
        End If
        If pokedata(3) <> 0 Then
            Command = New System.Data.SqlServerCe.SqlCeCommand("SELECT damage_factor FROM type_efficacy WHERE target_type_id = " & pokedata(3) & " ORDER BY damage_type_id ASC", Connection)
            myReader = Command.ExecuteReader(CommandBehavior.Default)
            iteration = 0
            While myReader.Read()
                efficacy2(iteration) = myReader.GetByte(0)
                iteration += 1
            End While
            myReader.Close()

            For iteration = 0 To 16
                efficacyNormalized(iteration) = Int((Val(efficacy1(iteration))) * (Val(efficacy2(iteration)) / 100))
            Next
            pokedata(24) = efficacyNormalized
        Else
            pokedata(24) = efficacy1
        End If



        Connection.Close()
        Return pokedata
    End Function

    Public Sub GetNature(ByRef SelectedIndex As UShort, ByRef mult As Double())
        mult = {1.0, 1.0, 1.0, 1.0, 1.0}
        If (SelectedIndex > 0) Then
            Try
                Connection.Open()
            Catch ex As Exception
                MsgBox("Could not connect to the database '" & DBFilename & "'.", MsgBoxStyle.Critical)
                Exit Sub
            End Try

            Dim Command As New System.Data.SqlServerCe.SqlCeCommand("SELECT decreased_stat_id, increased_stat_id FROM natures WHERE natures.nature_id = " & SelectedIndex, Connection)
            Dim myReader As System.Data.SqlServerCe.SqlCeDataReader = Command.ExecuteReader(CommandBehavior.Default)

            While myReader.Read()
                If myReader.GetByte(0) <> myReader.GetByte(1) Then
                    mult(myReader.GetByte(0) - 2) = 0.9
                    mult(myReader.GetByte(1) - 2) = 1.1
                End If
            End While

            myReader.Close()
            Connection.Close()
        End If
        CalcStats1()
        CalcStats2()
        CalcStats3()
        CalcStats4()
        CalcStats5()
        CalcStats6()
    End Sub

    Public Function GetMove(ByRef MoveID As UShort)
        Dim movedata() As Byte = {0, 0, 0, 0, 0}

        Try
            Connection.Open()
        Catch ex As Exception
            MsgBox("Could not connect to the database '" & DBFilename & "'.", MsgBoxStyle.Critical)
            Return movedata
        End Try

        Dim Command As New System.Data.SqlServerCe.SqlCeCommand("SELECT type_id, power, pp, accuracy, damage_class_id FROM moves WHERE moves.id = " & MoveID, Connection)
        Dim myReader As System.Data.SqlServerCe.SqlCeDataReader = Command.ExecuteReader(CommandBehavior.Default)

        While myReader.Read()
            movedata(0) = myReader.GetByte(0)
            If Not myReader.IsDBNull(1) Then movedata(1) = myReader.GetByte(1)
            movedata(2) = myReader.GetByte(2)
            If Not myReader.IsDBNull(3) Then movedata(3) = myReader.GetByte(3)
            movedata(4) = myReader.GetByte(4)
        End While

        myReader.Close()
        Connection.Close()

        Return movedata
    End Function

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        Dim A As New Form1()
        A.Show()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If MsgBox("Are you sure you want to close this window?" & vbCrLf & vbCrLf & "Any unsaved data will be lost.", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            e.Cancel = True
        End If
    End Sub

    Private Sub pkm1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm1.SelectedIndexChanged
        Dim pokedata() As Object = UpdateMon(pkm1.SelectedIndex)

        pkm1_selected = pokedata(0)
        pkm1_genus.Text = pokedata(1)
        pkm1_height.Text = pokedata(22)
        pkm1_weight.Text = pokedata(23)
        pkm1_image.Image = My.Resources.ResourceManager.GetObject(pokedata(4))
        pkm1_type1.Image = My.Resources.ResourceManager.GetObject("type" & pokedata(2))
        pkm1_type2.Image = My.Resources.ResourceManager.GetObject("type" & pokedata(3))

        pkm1_hp.Text = pokedata(7)
        pkm1_hp.Width = pokedata(8)
        pkm1_attack.Text = pokedata(9)
        pkm1_attack.Width = pokedata(10)
        pkm1_defense.Text = pokedata(11)
        pkm1_defense.Width = pokedata(12)
        pkm1_specialattack.Text = pokedata(13)
        pkm1_specialattack.Width = pokedata(14)
        pkm1_specialdefense.Text = pokedata(15)
        pkm1_specialdefense.Width = pokedata(16)
        pkm1_speed.Text = pokedata(17)
        pkm1_speed.Width = pokedata(18)
        pkm1_total.Text = pokedata(19)

        pkm1_ability.Items.Clear()
        For Each Item In pokedata(20)
            If Not Item Is Nothing Then pkm1_ability.Items.Add(Item)
        Next
        If pkm1_ability.Items.Count > 0 Then
            pkm1_ability.Enabled = True
            pkm1_ability.SelectedIndex = 0
        Else
            pkm1_ability.Enabled = False
        End If

        If pkm1_selected > 0 Then
            pkm1_nature.Enabled = True
            pkm1_iv1.Enabled = True
            pkm1_iv2.Enabled = True
            pkm1_iv3.Enabled = True
            pkm1_iv4.Enabled = True
            pkm1_iv5.Enabled = True
            pkm1_iv6.Enabled = True
            pkm1_ev1.Enabled = True
            pkm1_ev2.Enabled = True
            pkm1_ev3.Enabled = True
            pkm1_ev4.Enabled = True
            pkm1_ev5.Enabled = True
            pkm1_ev6.Enabled = True
            pkm1_level.Enabled = True
            pkm1_helditem.Enabled = True
            pkm1_move1.Enabled = True
            pkm1_move2.Enabled = True
            pkm1_move3.Enabled = True
            pkm1_move4.Enabled = True
            pkm1_nickname.Enabled = True
            pkm1_helditem_image.Visible = True
            pkm1_shiny.Enabled = True
            pkm1_gender.Enabled = True
        Else
            pkm1_nature.Enabled = False
            pkm1_iv1.Enabled = False
            pkm1_iv2.Enabled = False
            pkm1_iv3.Enabled = False
            pkm1_iv4.Enabled = False
            pkm1_iv5.Enabled = False
            pkm1_iv6.Enabled = False
            pkm1_ev1.Enabled = False
            pkm1_ev2.Enabled = False
            pkm1_ev3.Enabled = False
            pkm1_ev4.Enabled = False
            pkm1_ev5.Enabled = False
            pkm1_ev6.Enabled = False
            pkm1_level.Enabled = False
            pkm1_helditem.Enabled = False
            pkm1_move1.Enabled = False
            pkm1_move2.Enabled = False
            pkm1_move3.Enabled = False
            pkm1_move4.Enabled = False
            pkm1_nickname.Enabled = False
            pkm1_helditem_image.Visible = False
            pkm1_shiny.Enabled = False
            pkm1_gender.Enabled = False
        End If


        For i = 1 To pkm1_weaknesses.Controls.Count
            pkm1_weaknesses.Controls(0).Dispose()
        Next
        For i = 1 To pkm1_resistances.Controls.Count
            pkm1_resistances.Controls(0).Dispose()
        Next

        Dim StartWeak As UShort = 4
        Dim StartResist As UShort = 4

        For i = 0 To 16
            If pokedata(24)(i) > 100 Then
                Dim NewWeakness As New PictureBox
                NewWeakness.Size = New System.Drawing.Size(36, 18)
                NewWeakness.Parent = pkm1_weaknesses
                NewWeakness.SizeMode = PictureBoxSizeMode.CenterImage
                NewWeakness.Image = My.Resources.ResourceManager.GetObject("type" & (i + 1))
                NewWeakness.Tag = "type" & (i + 1)
                NewWeakness.Location = New System.Drawing.Point(StartWeak, 12)
                '4x
                If pokedata(24)(i) > 200 Then
                    NewWeakness.BackColor = Color.DarkRed
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewWeakness, "4×")
                Else
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewWeakness, "2×")
                End If
                StartWeak += 38
            ElseIf pokedata(24)(i) < 100 Then
                Dim NewResistance As New PictureBox
                NewResistance.Size = New System.Drawing.Size(36, 18)
                NewResistance.Parent = pkm1_resistances
                NewResistance.SizeMode = PictureBoxSizeMode.CenterImage
                NewResistance.Image = My.Resources.ResourceManager.GetObject("type" & (i + 1))
                NewResistance.Tag = "type" & (i + 1)
                NewResistance.Location = New System.Drawing.Point(StartResist, 12)
                '1/4x or 0x
                If pokedata(24)(i) = 0 Then
                    NewResistance.BackColor = Color.Black
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "0×")
                ElseIf pokedata(24)(i) < 50 Then
                    NewResistance.BackColor = Color.Green
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "¼×")
                Else
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "½×")
                End If
                StartResist += 38
            End If
        Next

        pkm1_species.Text = pokedata(21)

        pokemonNames(0) = pokedata(5)
        UpdateTab1()

        CalcStats1()

    End Sub

    Public Sub UpdateTab1()
        If pkm1.SelectedIndex > 0 Then
            If pkm1_nickname.Text.Length > 0 Then
                TabPage1.Text = "1. " & pkm1_nickname.Text
            Else
                TabPage1.Text = "1. " & pokemonNames(0)
            End If
        Else
            TabPage1.Text = "Pokémon 1"
        End If
    End Sub

    Private Sub pkm1_iv1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm1_iv1.ValueChanged, pkm1_iv2.ValueChanged, pkm1_iv3.ValueChanged, pkm1_iv4.ValueChanged, pkm1_iv5.ValueChanged, pkm1_iv6.ValueChanged
        pkm1_ivtotal.Text = pkm1_iv1.Value + pkm1_iv2.Value + _
            pkm1_iv3.Value + pkm1_iv4.Value + _
            pkm1_iv5.Value + pkm1_iv6.Value

        CalcStats1()
    End Sub

    Private Sub pkm1_ev1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm1_ev1.ValueChanged, pkm1_ev2.ValueChanged, pkm1_ev3.ValueChanged, pkm1_ev4.ValueChanged, pkm1_ev5.ValueChanged, pkm1_ev6.ValueChanged
        Dim sum As UShort
        sum = pkm1_ev1.Value + pkm1_ev2.Value + _
            pkm1_ev3.Value + pkm1_ev4.Value + _
            pkm1_ev5.Value + pkm1_ev6.Value

        pkm1_evtotal.Text = sum
        If sum > 510 Then
            pkm1_evtotal.ForeColor = Color.Red
        Else
            pkm1_evtotal.ForeColor = Color.Black
        End If

        CalcStats1()
    End Sub

    Public Sub CalcStats1()
        Dim nature As Double = 1.0
        If Val(pkm1_hp.Text) = 0 Then
            pkm1_stat1.Text = 0
            pkm1_stat2.Text = 0
            pkm1_stat3.Text = 0
            pkm1_stat4.Text = 0
            pkm1_stat5.Text = 0
            pkm1_stat6.Text = 0

        Else
            If Val(pkm1_hp.Text) = 1 Then
                pkm1_stat1.Text = 1
            Else
                pkm1_stat1.Text = Math.Floor((pkm1_iv1.Value + (2 * Val(pkm1_hp.Text)) + pkm1_ev1.Value / 4.0 + 100.0) * pkm1_level.Value / 100.0) + 10
            End If
            pkm1_stat2.Text = Math.Floor(Math.Floor((pkm1_iv2.Value + (2 * Val(pkm1_attack.Text)) + pkm1_ev2.Value / 4.0) * pkm1_level.Value / 100.0 + 5) * pkm1_nature_multipliers(0))
            pkm1_stat3.Text = Math.Floor(Math.Floor((pkm1_iv3.Value + (2 * Val(pkm1_defense.Text)) + pkm1_ev3.Value / 4.0) * pkm1_level.Value / 100.0 + 5) * pkm1_nature_multipliers(1))
            pkm1_stat4.Text = Math.Floor(Math.Floor((pkm1_iv4.Value + (2 * Val(pkm1_specialattack.Text)) + pkm1_ev4.Value / 4.0) * pkm1_level.Value / 100.0 + 5) * pkm1_nature_multipliers(2))
            pkm1_stat5.Text = Math.Floor(Math.Floor((pkm1_iv5.Value + (2 * Val(pkm1_specialdefense.Text)) + pkm1_ev5.Value / 4.0) * pkm1_level.Value / 100.0 + 5) * pkm1_nature_multipliers(3))
            pkm1_stat6.Text = Math.Floor(Math.Floor((pkm1_iv6.Value + (2 * Val(pkm1_speed.Text)) + pkm1_ev6.Value / 4.0) * pkm1_level.Value / 100.0 + 5) * pkm1_nature_multipliers(4))

        End If

        'Set colors based on nature
        pkm1_stat2.ForeColor = Color.Black
        pkm1_stat3.ForeColor = Color.Black
        pkm1_stat4.ForeColor = Color.Black
        pkm1_stat5.ForeColor = Color.Black
        pkm1_stat6.ForeColor = Color.Black

        If pkm1_nature_multipliers(0) > 1.0 Then
            pkm1_stat2.ForeColor = Color.DarkRed
        ElseIf pkm1_nature_multipliers(1) > 1.0 Then
            pkm1_stat3.ForeColor = Color.DarkRed
        ElseIf pkm1_nature_multipliers(2) > 1.0 Then
            pkm1_stat4.ForeColor = Color.DarkRed
        ElseIf pkm1_nature_multipliers(3) > 1.0 Then
            pkm1_stat5.ForeColor = Color.DarkRed
        ElseIf pkm1_nature_multipliers(4) > 1.0 Then
            pkm1_stat6.ForeColor = Color.DarkRed
        End If
        If pkm1_nature_multipliers(0) < 1.0 Then
            pkm1_stat2.ForeColor = Color.DarkBlue
        ElseIf pkm1_nature_multipliers(1) < 1.0 Then
            pkm1_stat3.ForeColor = Color.DarkBlue
        ElseIf pkm1_nature_multipliers(2) < 1.0 Then
            pkm1_stat4.ForeColor = Color.DarkBlue
        ElseIf pkm1_nature_multipliers(3) < 1.0 Then
            pkm1_stat5.ForeColor = Color.DarkBlue
        ElseIf pkm1_nature_multipliers(4) < 1.0 Then
            pkm1_stat6.ForeColor = Color.DarkBlue
        End If


        'Calculate Hidden Power
        If Val(pkm1_hp.Text) = 0 Then
            pkm1_hiddenpower.Image = Nothing
            pkm1_hiddenpower_base.Text = ""
        Else
            Dim HiddenPower As Double = _
                (((pkm1_iv1.Value Mod 2) + 2 * (pkm1_iv2.Value Mod 2) + _
                 4 * (pkm1_iv3.Value Mod 2) + 8 * (pkm1_iv6.Value Mod 2) + _
                 16 * (pkm1_iv4.Value Mod 2) + 32 * (pkm1_iv5.Value Mod 2)) * _
                 15) / 63.0
            Dim HiddenPowerBase As Double = _
                ((Math.Floor((pkm1_iv1.Value Mod 4) / 2) + 2 * Math.Floor((pkm1_iv2.Value Mod 4) / 2) + _
                 4 * Math.Floor((pkm1_iv3.Value Mod 4) / 2) + 8 * Math.Floor((pkm1_iv6.Value Mod 4) / 2) + _
                 16 * Math.Floor((pkm1_iv4.Value Mod 4) / 2) + 32 * Math.Floor((pkm1_iv5.Value Mod 4) / 2)) * _
                 40) / 63.0 + 30

            pkm1_hiddenpower.Image = My.Resources.ResourceManager.GetObject("type" & (Math.Floor(HiddenPower) + 2))
            pkm1_hiddenpower_base.Text = Math.Floor(HiddenPowerBase)
        End If

    End Sub

    Private Sub pkm1_level_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm1_level.ValueChanged
        CalcStats1()
    End Sub

    Private Sub pkm1_nature_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm1_nature.SelectedIndexChanged
        GetNature(pkm1_nature.SelectedIndex, pkm1_nature_multipliers)
    End Sub

    Private Sub pkm1_move1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm1_move1.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm1_move1.SelectedIndex)

        pkm1_move1_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm1_move1_power.Text = "---"
        Else
            pkm1_move1_power.Text = movedata(1)
        End If
        pkm1_move1_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm1_move1_accuracy.Text = "---"
        Else
            pkm1_move1_accuracy.Text = movedata(3)
        End If
        pkm1_move1_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm1_move2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm1_move2.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm1_move2.SelectedIndex)

        pkm1_move2_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm1_move2_power.Text = "---"
        Else
            pkm1_move2_power.Text = movedata(1)
        End If
        pkm1_move2_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm1_move2_accuracy.Text = "---"
        Else
            pkm1_move2_accuracy.Text = movedata(3)
        End If
        pkm1_move2_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm1_move3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm1_move3.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm1_move3.SelectedIndex)

        pkm1_move3_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm1_move3_power.Text = "---"
        Else
            pkm1_move3_power.Text = movedata(1)
        End If
        pkm1_move3_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm1_move3_accuracy.Text = "---"
        Else
            pkm1_move3_accuracy.Text = movedata(3)
        End If
        pkm1_move3_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm1_move4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm1_move4.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm1_move4.SelectedIndex)

        pkm1_move4_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm1_move4_power.Text = "---"
        Else
            pkm1_move4_power.Text = movedata(1)
        End If
        pkm1_move4_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm1_move4_accuracy.Text = "---"
        Else
            pkm1_move4_accuracy.Text = movedata(3)
        End If
        pkm1_move4_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm1_nickname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm1_nickname.TextChanged
        UpdateTab1()
    End Sub

    Private Sub pkm1_helditem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm1_helditem.SelectedIndexChanged
        Try
            Connection.Open()
        Catch ex As Exception
            MsgBox("Could not connect to the database '" & DBFilename & "'.", MsgBoxStyle.Critical)
            Exit Sub
        End Try

        Dim Command As New System.Data.SqlServerCe.SqlCeCommand("SELECT identifier FROM items WHERE items.id = " & pkm1_helditem.SelectedIndex, Connection)
        Dim myReader As System.Data.SqlServerCe.SqlCeDataReader = Command.ExecuteReader(CommandBehavior.Default)

        pkm1_helditem_image.Image = Nothing
        While myReader.Read()
            pkm1_helditem_image.Image = My.Resources.ResourceManager.GetObject(myReader.GetString(0).Replace("-", "_"))
        End While

        myReader.Close()
        Connection.Close()
    End Sub

    'Private Sub pkm1_find_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    For Item = 0 To pkm1.Items.Count - 1
    '        If pkm1.Items(Item).ToString.ToLower.Contains(pkm1_find.Text.ToLower) Then
    '            pkm1.SelectedIndex = Item
    '            Exit Sub
    '        End If
    '    Next
    'End Sub

    Private Sub pkm2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm2.SelectedIndexChanged
        Dim pokedata() As Object = UpdateMon(pkm2.SelectedIndex)

        pkm2_selected = pokedata(0)
        pkm2_genus.Text = pokedata(1)
        pkm2_height.Text = pokedata(22)
        pkm2_weight.Text = pokedata(23)
        pkm2_image.Image = My.Resources.ResourceManager.GetObject(pokedata(4))
        pkm2_type1.Image = My.Resources.ResourceManager.GetObject("type" & pokedata(2))
        pkm2_type2.Image = My.Resources.ResourceManager.GetObject("type" & pokedata(3))

        pkm2_hp.Text = pokedata(7)
        pkm2_hp.Width = pokedata(8)
        pkm2_attack.Text = pokedata(9)
        pkm2_attack.Width = pokedata(10)
        pkm2_defense.Text = pokedata(11)
        pkm2_defense.Width = pokedata(12)
        pkm2_specialattack.Text = pokedata(13)
        pkm2_specialattack.Width = pokedata(14)
        pkm2_specialdefense.Text = pokedata(15)
        pkm2_specialdefense.Width = pokedata(16)
        pkm2_speed.Text = pokedata(17)
        pkm2_speed.Width = pokedata(18)
        pkm2_total.Text = pokedata(19)

        pkm2_ability.Items.Clear()
        For Each Item In pokedata(20)
            If Not Item Is Nothing Then pkm2_ability.Items.Add(Item)
        Next
        If pkm2_ability.Items.Count > 0 Then
            pkm2_ability.Enabled = True
            pkm2_ability.SelectedIndex = 0
        Else
            pkm2_ability.Enabled = False
        End If

        If pkm2_selected > 0 Then
            pkm2_nature.Enabled = True
            pkm2_iv1.Enabled = True
            pkm2_iv2.Enabled = True
            pkm2_iv3.Enabled = True
            pkm2_iv4.Enabled = True
            pkm2_iv5.Enabled = True
            pkm2_iv6.Enabled = True
            pkm2_ev1.Enabled = True
            pkm2_ev2.Enabled = True
            pkm2_ev3.Enabled = True
            pkm2_ev4.Enabled = True
            pkm2_ev5.Enabled = True
            pkm2_ev6.Enabled = True
            pkm2_level.Enabled = True
            pkm2_helditem.Enabled = True
            pkm2_move1.Enabled = True
            pkm2_move2.Enabled = True
            pkm2_move3.Enabled = True
            pkm2_move4.Enabled = True
            pkm2_nickname.Enabled = True
            pkm2_helditem_image.Visible = True
            pkm2_shiny.Enabled = True
            pkm2_gender.Enabled = True
        Else
            pkm2_nature.Enabled = False
            pkm2_iv1.Enabled = False
            pkm2_iv2.Enabled = False
            pkm2_iv3.Enabled = False
            pkm2_iv4.Enabled = False
            pkm2_iv5.Enabled = False
            pkm2_iv6.Enabled = False
            pkm2_ev1.Enabled = False
            pkm2_ev2.Enabled = False
            pkm2_ev3.Enabled = False
            pkm2_ev4.Enabled = False
            pkm2_ev5.Enabled = False
            pkm2_ev6.Enabled = False
            pkm2_level.Enabled = False
            pkm2_helditem.Enabled = False
            pkm2_move1.Enabled = False
            pkm2_move2.Enabled = False
            pkm2_move3.Enabled = False
            pkm2_move4.Enabled = False
            pkm2_nickname.Enabled = False
            pkm2_helditem_image.Visible = False
            pkm2_shiny.Enabled = False
            pkm2_gender.Enabled = False
        End If


        For i = 1 To pkm2_weaknesses.Controls.Count
            pkm2_weaknesses.Controls(0).Dispose()
        Next
        For i = 1 To pkm2_resistances.Controls.Count
            pkm2_resistances.Controls(0).Dispose()
        Next

        Dim StartWeak As UShort = 4
        Dim StartResist As UShort = 4

        For i = 0 To 16
            If pokedata(24)(i) > 100 Then
                Dim NewWeakness As New PictureBox
                NewWeakness.Size = New System.Drawing.Size(36, 18)
                NewWeakness.Parent = pkm2_weaknesses
                NewWeakness.SizeMode = PictureBoxSizeMode.CenterImage
                NewWeakness.Image = My.Resources.ResourceManager.GetObject("type" & (i + 1))
                NewWeakness.Tag = "type" & (i + 1)
                NewWeakness.Location = New System.Drawing.Point(StartWeak, 12)
                '4x
                If pokedata(24)(i) > 200 Then
                    NewWeakness.BackColor = Color.DarkRed
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewWeakness, "4×")
                Else
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewWeakness, "2×")
                End If
                StartWeak += 38
            ElseIf pokedata(24)(i) < 100 Then
                Dim NewResistance As New PictureBox
                NewResistance.Size = New System.Drawing.Size(36, 18)
                NewResistance.Parent = pkm2_resistances
                NewResistance.SizeMode = PictureBoxSizeMode.CenterImage
                NewResistance.Image = My.Resources.ResourceManager.GetObject("type" & (i + 1))
                NewResistance.Tag = "type" & (i + 1)
                NewResistance.Location = New System.Drawing.Point(StartResist, 12)
                '1/4x or 0x
                If pokedata(24)(i) = 0 Then
                    NewResistance.BackColor = Color.Black
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "0×")
                ElseIf pokedata(24)(i) < 50 Then
                    NewResistance.BackColor = Color.Green
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "¼×")
                Else
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "½×")
                End If
                StartResist += 38
            End If
        Next

        pkm2_species.Text = pokedata(21)

        pokemonNames(1) = pokedata(5)
        UpdateTab2()

        CalcStats2()
    End Sub

    Public Sub UpdateTab2()
        If pkm2.SelectedIndex > 0 Then
            If pkm2_nickname.Text.Length > 0 Then
                TabPage2.Text = "2. " & pkm2_nickname.Text
            Else
                TabPage2.Text = "2. " & pokemonNames(1)
            End If
        Else
            TabPage2.Text = "Pokémon 2"
        End If
    End Sub

    Private Sub pkm2_iv1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm2_iv1.ValueChanged, pkm2_iv2.ValueChanged, pkm2_iv3.ValueChanged, pkm2_iv4.ValueChanged, pkm2_iv5.ValueChanged, pkm2_iv6.ValueChanged
        pkm2_ivtotal.Text = pkm2_iv1.Value + pkm2_iv2.Value + _
            pkm2_iv3.Value + pkm2_iv4.Value + _
            pkm2_iv5.Value + pkm2_iv6.Value

        CalcStats2()
    End Sub

    Private Sub pkm2_ev1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm2_ev1.ValueChanged, pkm2_ev2.ValueChanged, pkm2_ev3.ValueChanged, pkm2_ev4.ValueChanged, pkm2_ev5.ValueChanged, pkm2_ev6.ValueChanged
        Dim sum As UShort
        sum = pkm2_ev1.Value + pkm2_ev2.Value + _
            pkm2_ev3.Value + pkm2_ev4.Value + _
            pkm2_ev5.Value + pkm2_ev6.Value

        pkm2_evtotal.Text = sum
        If sum > 510 Then
            pkm2_evtotal.ForeColor = Color.Red
        Else
            pkm2_evtotal.ForeColor = Color.Black
        End If

        CalcStats2()
    End Sub

    Public Sub CalcStats2()
        Dim nature As Double = 1.0
        If Val(pkm2_hp.Text) = 0 Then
            pkm2_stat1.Text = 0
            pkm2_stat2.Text = 0
            pkm2_stat3.Text = 0
            pkm2_stat4.Text = 0
            pkm2_stat5.Text = 0
            pkm2_stat6.Text = 0

        Else
            If Val(pkm2_hp.Text) = 1 Then
                pkm2_stat1.Text = 1
            Else
                pkm2_stat1.Text = Math.Floor((pkm2_iv1.Value + (2 * Val(pkm2_hp.Text)) + pkm2_ev1.Value / 4.0 + 100.0) * pkm2_level.Value / 100.0) + 10
            End If
            pkm2_stat2.Text = Math.Floor(Math.Floor((pkm2_iv2.Value + (2 * Val(pkm2_attack.Text)) + pkm2_ev2.Value / 4.0) * pkm2_level.Value / 100.0 + 5) * pkm2_nature_multipliers(0))
            pkm2_stat3.Text = Math.Floor(Math.Floor((pkm2_iv3.Value + (2 * Val(pkm2_defense.Text)) + pkm2_ev3.Value / 4.0) * pkm2_level.Value / 100.0 + 5) * pkm2_nature_multipliers(1))
            pkm2_stat4.Text = Math.Floor(Math.Floor((pkm2_iv4.Value + (2 * Val(pkm2_specialattack.Text)) + pkm2_ev4.Value / 4.0) * pkm2_level.Value / 100.0 + 5) * pkm2_nature_multipliers(2))
            pkm2_stat5.Text = Math.Floor(Math.Floor((pkm2_iv5.Value + (2 * Val(pkm2_specialdefense.Text)) + pkm2_ev5.Value / 4.0) * pkm2_level.Value / 100.0 + 5) * pkm2_nature_multipliers(3))
            pkm2_stat6.Text = Math.Floor(Math.Floor((pkm2_iv6.Value + (2 * Val(pkm2_speed.Text)) + pkm2_ev6.Value / 4.0) * pkm2_level.Value / 100.0 + 5) * pkm2_nature_multipliers(4))

        End If

        'Set colors based on nature
        pkm2_stat2.ForeColor = Color.Black
        pkm2_stat3.ForeColor = Color.Black
        pkm2_stat4.ForeColor = Color.Black
        pkm2_stat5.ForeColor = Color.Black
        pkm2_stat6.ForeColor = Color.Black

        If pkm2_nature_multipliers(0) > 1.0 Then
            pkm2_stat2.ForeColor = Color.DarkRed
        ElseIf pkm2_nature_multipliers(1) > 1.0 Then
            pkm2_stat3.ForeColor = Color.DarkRed
        ElseIf pkm2_nature_multipliers(2) > 1.0 Then
            pkm2_stat4.ForeColor = Color.DarkRed
        ElseIf pkm2_nature_multipliers(3) > 1.0 Then
            pkm2_stat5.ForeColor = Color.DarkRed
        ElseIf pkm2_nature_multipliers(4) > 1.0 Then
            pkm2_stat6.ForeColor = Color.DarkRed
        End If
        If pkm2_nature_multipliers(0) < 1.0 Then
            pkm2_stat2.ForeColor = Color.DarkBlue
        ElseIf pkm2_nature_multipliers(1) < 1.0 Then
            pkm2_stat3.ForeColor = Color.DarkBlue
        ElseIf pkm2_nature_multipliers(2) < 1.0 Then
            pkm2_stat4.ForeColor = Color.DarkBlue
        ElseIf pkm2_nature_multipliers(3) < 1.0 Then
            pkm2_stat5.ForeColor = Color.DarkBlue
        ElseIf pkm2_nature_multipliers(4) < 1.0 Then
            pkm2_stat6.ForeColor = Color.DarkBlue
        End If


        'Calculate Hidden Power
        If Val(pkm2_hp.Text) = 0 Then
            pkm2_hiddenpower.Image = Nothing
            pkm2_hiddenpower_base.Text = ""
        Else
            Dim HiddenPower As Double = _
                (((pkm2_iv1.Value Mod 2) + 2 * (pkm2_iv2.Value Mod 2) + _
                 4 * (pkm2_iv3.Value Mod 2) + 8 * (pkm2_iv6.Value Mod 2) + _
                 16 * (pkm2_iv4.Value Mod 2) + 32 * (pkm2_iv5.Value Mod 2)) * _
                 15) / 63.0
            Dim HiddenPowerBase As Double = _
                ((Math.Floor((pkm2_iv1.Value Mod 4) / 2) + 2 * Math.Floor((pkm2_iv2.Value Mod 4) / 2) + _
                 4 * Math.Floor((pkm2_iv3.Value Mod 4) / 2) + 8 * Math.Floor((pkm2_iv6.Value Mod 4) / 2) + _
                 16 * Math.Floor((pkm2_iv4.Value Mod 4) / 2) + 32 * Math.Floor((pkm2_iv5.Value Mod 4) / 2)) * _
                 40) / 63.0 + 30

            pkm2_hiddenpower.Image = My.Resources.ResourceManager.GetObject("type" & (Math.Floor(HiddenPower) + 2))
            pkm2_hiddenpower_base.Text = Math.Floor(HiddenPowerBase)
        End If

    End Sub

    Private Sub pkm2_level_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm2_level.ValueChanged
        CalcStats2()
    End Sub

    Private Sub pkm2_nature_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm2_nature.SelectedIndexChanged
        GetNature(pkm2_nature.SelectedIndex, pkm2_nature_multipliers)
    End Sub

    Private Sub pkm2_move1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm2_move1.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm2_move1.SelectedIndex)

        pkm2_move1_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm2_move1_power.Text = "---"
        Else
            pkm2_move1_power.Text = movedata(1)
        End If
        pkm2_move1_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm2_move1_accuracy.Text = "---"
        Else
            pkm2_move1_accuracy.Text = movedata(3)
        End If
        pkm2_move1_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm2_move2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm2_move2.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm2_move2.SelectedIndex)

        pkm2_move2_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm2_move2_power.Text = "---"
        Else
            pkm2_move2_power.Text = movedata(1)
        End If
        pkm2_move2_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm2_move2_accuracy.Text = "---"
        Else
            pkm2_move2_accuracy.Text = movedata(3)
        End If
        pkm2_move2_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm2_move3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm2_move3.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm2_move3.SelectedIndex)

        pkm2_move3_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm2_move3_power.Text = "---"
        Else
            pkm2_move3_power.Text = movedata(1)
        End If
        pkm2_move3_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm2_move3_accuracy.Text = "---"
        Else
            pkm2_move3_accuracy.Text = movedata(3)
        End If
        pkm2_move3_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm2_move4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm2_move4.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm2_move4.SelectedIndex)

        pkm2_move4_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm2_move4_power.Text = "---"
        Else
            pkm2_move4_power.Text = movedata(1)
        End If
        pkm2_move4_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm2_move4_accuracy.Text = "---"
        Else
            pkm2_move4_accuracy.Text = movedata(3)
        End If
        pkm2_move4_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm2_nickname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm2_nickname.TextChanged
        UpdateTab2()
    End Sub

    Private Sub pkm2_helditem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm2_helditem.SelectedIndexChanged
        Try
            Connection.Open()
        Catch ex As Exception
            MsgBox("Could not connect to the database '" & DBFilename & "'.", MsgBoxStyle.Critical)
            Exit Sub
        End Try

        Dim Command As New System.Data.SqlServerCe.SqlCeCommand("SELECT identifier FROM items WHERE items.id = " & pkm2_helditem.SelectedIndex, Connection)
        Dim myReader As System.Data.SqlServerCe.SqlCeDataReader = Command.ExecuteReader(CommandBehavior.Default)

        pkm2_helditem_image.Image = Nothing
        While myReader.Read()
            pkm2_helditem_image.Image = My.Resources.ResourceManager.GetObject(myReader.GetString(0).Replace("-", "_"))
        End While

        myReader.Close()
        Connection.Close()
    End Sub

    'Private Sub pkm2_find_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    For Item = 0 To pkm2.Items.Count - 1
    '        If pkm2.Items(Item).ToString.ToLower.Contains(pkm2_find.Text.ToLower) Then
    '            pkm2.SelectedIndex = Item
    '            Exit Sub
    '        End If
    '    Next
    'End Sub

    Private Sub pkm3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm3.SelectedIndexChanged
        Dim pokedata() As Object = UpdateMon(pkm3.SelectedIndex)

        pkm3_selected = pokedata(0)
        pkm3_genus.Text = pokedata(1)
        pkm3_height.Text = pokedata(22)
        pkm3_weight.Text = pokedata(23)
        pkm3_image.Image = My.Resources.ResourceManager.GetObject(pokedata(4))
        pkm3_type1.Image = My.Resources.ResourceManager.GetObject("type" & pokedata(2))
        pkm3_type2.Image = My.Resources.ResourceManager.GetObject("type" & pokedata(3))

        pkm3_hp.Text = pokedata(7)
        pkm3_hp.Width = pokedata(8)
        pkm3_attack.Text = pokedata(9)
        pkm3_attack.Width = pokedata(10)
        pkm3_defense.Text = pokedata(11)
        pkm3_defense.Width = pokedata(12)
        pkm3_specialattack.Text = pokedata(13)
        pkm3_specialattack.Width = pokedata(14)
        pkm3_specialdefense.Text = pokedata(15)
        pkm3_specialdefense.Width = pokedata(16)
        pkm3_speed.Text = pokedata(17)
        pkm3_speed.Width = pokedata(18)
        pkm3_total.Text = pokedata(19)

        pkm3_ability.Items.Clear()
        For Each Item In pokedata(20)
            If Not Item Is Nothing Then pkm3_ability.Items.Add(Item)
        Next
        If pkm3_ability.Items.Count > 0 Then
            pkm3_ability.Enabled = True
            pkm3_ability.SelectedIndex = 0
        Else
            pkm3_ability.Enabled = False
        End If

        If pkm3_selected > 0 Then
            pkm3_nature.Enabled = True
            pkm3_iv1.Enabled = True
            pkm3_iv2.Enabled = True
            pkm3_iv3.Enabled = True
            pkm3_iv4.Enabled = True
            pkm3_iv5.Enabled = True
            pkm3_iv6.Enabled = True
            pkm3_ev1.Enabled = True
            pkm3_ev2.Enabled = True
            pkm3_ev3.Enabled = True
            pkm3_ev4.Enabled = True
            pkm3_ev5.Enabled = True
            pkm3_ev6.Enabled = True
            pkm3_level.Enabled = True
            pkm3_helditem.Enabled = True
            pkm3_move1.Enabled = True
            pkm3_move2.Enabled = True
            pkm3_move3.Enabled = True
            pkm3_move4.Enabled = True
            pkm3_nickname.Enabled = True
            pkm3_helditem_image.Visible = True
            pkm3_shiny.Enabled = True
            pkm3_gender.Enabled = True
        Else
            pkm3_nature.Enabled = False
            pkm3_iv1.Enabled = False
            pkm3_iv2.Enabled = False
            pkm3_iv3.Enabled = False
            pkm3_iv4.Enabled = False
            pkm3_iv5.Enabled = False
            pkm3_iv6.Enabled = False
            pkm3_ev1.Enabled = False
            pkm3_ev2.Enabled = False
            pkm3_ev3.Enabled = False
            pkm3_ev4.Enabled = False
            pkm3_ev5.Enabled = False
            pkm3_ev6.Enabled = False
            pkm3_level.Enabled = False
            pkm3_helditem.Enabled = False
            pkm3_move1.Enabled = False
            pkm3_move2.Enabled = False
            pkm3_move3.Enabled = False
            pkm3_move4.Enabled = False
            pkm3_nickname.Enabled = False
            pkm3_helditem_image.Visible = False
            pkm3_shiny.Enabled = False
            pkm3_gender.Enabled = False
        End If


        For i = 1 To pkm3_weaknesses.Controls.Count
            pkm3_weaknesses.Controls(0).Dispose()
        Next
        For i = 1 To pkm3_resistances.Controls.Count
            pkm3_resistances.Controls(0).Dispose()
        Next

        Dim StartWeak As UShort = 4
        Dim StartResist As UShort = 4

        For i = 0 To 16
            If pokedata(24)(i) > 100 Then
                Dim NewWeakness As New PictureBox
                NewWeakness.Size = New System.Drawing.Size(36, 18)
                NewWeakness.Parent = pkm3_weaknesses
                NewWeakness.SizeMode = PictureBoxSizeMode.CenterImage
                NewWeakness.Image = My.Resources.ResourceManager.GetObject("type" & (i + 1))
                NewWeakness.Tag = "type" & (i + 1)
                NewWeakness.Location = New System.Drawing.Point(StartWeak, 12)
                '4x
                If pokedata(24)(i) > 200 Then
                    NewWeakness.BackColor = Color.DarkRed
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewWeakness, "4×")
                Else
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewWeakness, "2×")
                End If
                StartWeak += 38
            ElseIf pokedata(24)(i) < 100 Then
                Dim NewResistance As New PictureBox
                NewResistance.Size = New System.Drawing.Size(36, 18)
                NewResistance.Parent = pkm3_resistances
                NewResistance.SizeMode = PictureBoxSizeMode.CenterImage
                NewResistance.Image = My.Resources.ResourceManager.GetObject("type" & (i + 1))
                NewResistance.Tag = "type" & (i + 1)
                NewResistance.Location = New System.Drawing.Point(StartResist, 12)
                '1/4x or 0x
                If pokedata(24)(i) = 0 Then
                    NewResistance.BackColor = Color.Black
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "0×")
                ElseIf pokedata(24)(i) < 50 Then
                    NewResistance.BackColor = Color.Green
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "¼×")
                Else
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "½×")
                End If
                StartResist += 38
            End If
        Next

        pkm3_species.Text = pokedata(21)

        pokemonNames(2) = pokedata(5)
        UpdateTab3()

        CalcStats3()
    End Sub

    Public Sub UpdateTab3()
        If pkm3.SelectedIndex > 0 Then
            If pkm3_nickname.Text.Length > 0 Then
                TabPage3.Text = "3. " & pkm3_nickname.Text
            Else
                TabPage3.Text = "3. " & pokemonNames(2)
            End If
        Else
            TabPage3.Text = "Pokémon 3"
        End If
    End Sub

    Private Sub pkm3_iv1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm3_iv1.ValueChanged, pkm3_iv2.ValueChanged, pkm3_iv3.ValueChanged, pkm3_iv4.ValueChanged, pkm3_iv5.ValueChanged, pkm3_iv6.ValueChanged
        pkm3_ivtotal.Text = pkm3_iv1.Value + pkm3_iv2.Value + _
            pkm3_iv3.Value + pkm3_iv4.Value + _
            pkm3_iv5.Value + pkm3_iv6.Value

        CalcStats3()
    End Sub

    Private Sub pkm3_ev1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm3_ev1.ValueChanged, pkm3_ev2.ValueChanged, pkm3_ev3.ValueChanged, pkm3_ev4.ValueChanged, pkm3_ev5.ValueChanged, pkm3_ev6.ValueChanged
        Dim sum As UShort
        sum = pkm3_ev1.Value + pkm3_ev2.Value + _
            pkm3_ev3.Value + pkm3_ev4.Value + _
            pkm3_ev5.Value + pkm3_ev6.Value

        pkm3_evtotal.Text = sum
        If sum > 510 Then
            pkm3_evtotal.ForeColor = Color.Red
        Else
            pkm3_evtotal.ForeColor = Color.Black
        End If

        CalcStats3()
    End Sub

    Public Sub CalcStats3()
        Dim nature As Double = 1.0
        If Val(pkm3_hp.Text) = 0 Then
            pkm3_stat1.Text = 0
            pkm3_stat2.Text = 0
            pkm3_stat3.Text = 0
            pkm3_stat4.Text = 0
            pkm3_stat5.Text = 0
            pkm3_stat6.Text = 0

        Else
            If Val(pkm3_hp.Text) = 1 Then
                pkm3_stat1.Text = 1
            Else
                pkm3_stat1.Text = Math.Floor((pkm3_iv1.Value + (2 * Val(pkm3_hp.Text)) + pkm3_ev1.Value / 4.0 + 100.0) * pkm3_level.Value / 100.0) + 10
            End If
            pkm3_stat2.Text = Math.Floor(Math.Floor((pkm3_iv2.Value + (2 * Val(pkm3_attack.Text)) + pkm3_ev2.Value / 4.0) * pkm3_level.Value / 100.0 + 5) * pkm3_nature_multipliers(0))
            pkm3_stat3.Text = Math.Floor(Math.Floor((pkm3_iv3.Value + (2 * Val(pkm3_defense.Text)) + pkm3_ev3.Value / 4.0) * pkm3_level.Value / 100.0 + 5) * pkm3_nature_multipliers(1))
            pkm3_stat4.Text = Math.Floor(Math.Floor((pkm3_iv4.Value + (2 * Val(pkm3_specialattack.Text)) + pkm3_ev4.Value / 4.0) * pkm3_level.Value / 100.0 + 5) * pkm3_nature_multipliers(2))
            pkm3_stat5.Text = Math.Floor(Math.Floor((pkm3_iv5.Value + (2 * Val(pkm3_specialdefense.Text)) + pkm3_ev5.Value / 4.0) * pkm3_level.Value / 100.0 + 5) * pkm3_nature_multipliers(3))
            pkm3_stat6.Text = Math.Floor(Math.Floor((pkm3_iv6.Value + (2 * Val(pkm3_speed.Text)) + pkm3_ev6.Value / 4.0) * pkm3_level.Value / 100.0 + 5) * pkm3_nature_multipliers(4))

        End If

        'Set colors based on nature
        pkm3_stat2.ForeColor = Color.Black
        pkm3_stat3.ForeColor = Color.Black
        pkm3_stat4.ForeColor = Color.Black
        pkm3_stat5.ForeColor = Color.Black
        pkm3_stat6.ForeColor = Color.Black

        If pkm3_nature_multipliers(0) > 1.0 Then
            pkm3_stat2.ForeColor = Color.DarkRed
        ElseIf pkm3_nature_multipliers(1) > 1.0 Then
            pkm3_stat3.ForeColor = Color.DarkRed
        ElseIf pkm3_nature_multipliers(2) > 1.0 Then
            pkm3_stat4.ForeColor = Color.DarkRed
        ElseIf pkm3_nature_multipliers(3) > 1.0 Then
            pkm3_stat5.ForeColor = Color.DarkRed
        ElseIf pkm3_nature_multipliers(4) > 1.0 Then
            pkm3_stat6.ForeColor = Color.DarkRed
        End If
        If pkm3_nature_multipliers(0) < 1.0 Then
            pkm3_stat2.ForeColor = Color.DarkBlue
        ElseIf pkm3_nature_multipliers(1) < 1.0 Then
            pkm3_stat3.ForeColor = Color.DarkBlue
        ElseIf pkm3_nature_multipliers(2) < 1.0 Then
            pkm3_stat4.ForeColor = Color.DarkBlue
        ElseIf pkm3_nature_multipliers(3) < 1.0 Then
            pkm3_stat5.ForeColor = Color.DarkBlue
        ElseIf pkm3_nature_multipliers(4) < 1.0 Then
            pkm3_stat6.ForeColor = Color.DarkBlue
        End If


        'Calculate Hidden Power
        If Val(pkm3_hp.Text) = 0 Then
            pkm3_hiddenpower.Image = Nothing
            pkm3_hiddenpower_base.Text = ""
        Else
            Dim HiddenPower As Double = _
                (((pkm3_iv1.Value Mod 2) + 2 * (pkm3_iv2.Value Mod 2) + _
                 4 * (pkm3_iv3.Value Mod 2) + 8 * (pkm3_iv6.Value Mod 2) + _
                 16 * (pkm3_iv4.Value Mod 2) + 32 * (pkm3_iv5.Value Mod 2)) * _
                 15) / 63.0
            Dim HiddenPowerBase As Double = _
                ((Math.Floor((pkm3_iv1.Value Mod 4) / 2) + 2 * Math.Floor((pkm3_iv2.Value Mod 4) / 2) + _
                 4 * Math.Floor((pkm3_iv3.Value Mod 4) / 2) + 8 * Math.Floor((pkm3_iv6.Value Mod 4) / 2) + _
                 16 * Math.Floor((pkm3_iv4.Value Mod 4) / 2) + 32 * Math.Floor((pkm3_iv5.Value Mod 4) / 2)) * _
                 40) / 63.0 + 30

            pkm3_hiddenpower.Image = My.Resources.ResourceManager.GetObject("type" & (Math.Floor(HiddenPower) + 2))
            pkm3_hiddenpower_base.Text = Math.Floor(HiddenPowerBase)
        End If

    End Sub

    Private Sub pkm3_level_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm3_level.ValueChanged
        CalcStats3()
    End Sub

    Private Sub pkm3_nature_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm3_nature.SelectedIndexChanged
        GetNature(pkm3_nature.SelectedIndex, pkm3_nature_multipliers)
    End Sub

    Private Sub pkm3_move1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm3_move1.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm3_move1.SelectedIndex)

        pkm3_move1_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm3_move1_power.Text = "---"
        Else
            pkm3_move1_power.Text = movedata(1)
        End If
        pkm3_move1_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm3_move1_accuracy.Text = "---"
        Else
            pkm3_move1_accuracy.Text = movedata(3)
        End If
        pkm3_move1_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm3_move2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm3_move2.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm3_move2.SelectedIndex)

        pkm3_move2_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm3_move2_power.Text = "---"
        Else
            pkm3_move2_power.Text = movedata(1)
        End If
        pkm3_move2_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm3_move2_accuracy.Text = "---"
        Else
            pkm3_move2_accuracy.Text = movedata(3)
        End If
        pkm3_move2_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm3_move3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm3_move3.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm3_move3.SelectedIndex)

        pkm3_move3_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm3_move3_power.Text = "---"
        Else
            pkm3_move3_power.Text = movedata(1)
        End If
        pkm3_move3_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm3_move3_accuracy.Text = "---"
        Else
            pkm3_move3_accuracy.Text = movedata(3)
        End If
        pkm3_move3_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm3_move4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm3_move4.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm3_move4.SelectedIndex)

        pkm3_move4_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm3_move4_power.Text = "---"
        Else
            pkm3_move4_power.Text = movedata(1)
        End If
        pkm3_move4_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm3_move4_accuracy.Text = "---"
        Else
            pkm3_move4_accuracy.Text = movedata(3)
        End If
        pkm3_move4_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm3_nickname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm3_nickname.TextChanged
        UpdateTab3()
    End Sub

    Private Sub pkm3_helditem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm3_helditem.SelectedIndexChanged
        Try
            Connection.Open()
        Catch ex As Exception
            MsgBox("Could not connect to the database '" & DBFilename & "'.", MsgBoxStyle.Critical)
            Exit Sub
        End Try

        Dim Command As New System.Data.SqlServerCe.SqlCeCommand("SELECT identifier FROM items WHERE items.id = " & pkm3_helditem.SelectedIndex, Connection)
        Dim myReader As System.Data.SqlServerCe.SqlCeDataReader = Command.ExecuteReader(CommandBehavior.Default)

        pkm3_helditem_image.Image = Nothing
        While myReader.Read()
            pkm3_helditem_image.Image = My.Resources.ResourceManager.GetObject(myReader.GetString(0).Replace("-", "_"))
        End While

        myReader.Close()
        Connection.Close()
    End Sub

    'Private Sub pkm3_find_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    For Item = 0 To pkm3.Items.Count - 1
    '        If pkm3.Items(Item).ToString.ToLower.Contains(pkm3_find.Text.ToLower) Then
    '            pkm3.SelectedIndex = Item
    '            Exit Sub
    '        End If
    '    Next
    'End Sub

    Private Sub pkm4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm4.SelectedIndexChanged
        Dim pokedata() As Object = UpdateMon(pkm4.SelectedIndex)

        pkm4_selected = pokedata(0)
        pkm4_genus.Text = pokedata(1)
        pkm4_height.Text = pokedata(22)
        pkm4_weight.Text = pokedata(23)
        pkm4_image.Image = My.Resources.ResourceManager.GetObject(pokedata(4))
        pkm4_type1.Image = My.Resources.ResourceManager.GetObject("type" & pokedata(2))
        pkm4_type2.Image = My.Resources.ResourceManager.GetObject("type" & pokedata(3))

        pkm4_hp.Text = pokedata(7)
        pkm4_hp.Width = pokedata(8)
        pkm4_attack.Text = pokedata(9)
        pkm4_attack.Width = pokedata(10)
        pkm4_defense.Text = pokedata(11)
        pkm4_defense.Width = pokedata(12)
        pkm4_specialattack.Text = pokedata(13)
        pkm4_specialattack.Width = pokedata(14)
        pkm4_specialdefense.Text = pokedata(15)
        pkm4_specialdefense.Width = pokedata(16)
        pkm4_speed.Text = pokedata(17)
        pkm4_speed.Width = pokedata(18)
        pkm4_total.Text = pokedata(19)

        pkm4_ability.Items.Clear()
        For Each Item In pokedata(20)
            If Not Item Is Nothing Then pkm4_ability.Items.Add(Item)
        Next
        If pkm4_ability.Items.Count > 0 Then
            pkm4_ability.Enabled = True
            pkm4_ability.SelectedIndex = 0
        Else
            pkm4_ability.Enabled = False
        End If

        If pkm4_selected > 0 Then
            pkm4_nature.Enabled = True
            pkm4_iv1.Enabled = True
            pkm4_iv2.Enabled = True
            pkm4_iv3.Enabled = True
            pkm4_iv4.Enabled = True
            pkm4_iv5.Enabled = True
            pkm4_iv6.Enabled = True
            pkm4_ev1.Enabled = True
            pkm4_ev2.Enabled = True
            pkm4_ev3.Enabled = True
            pkm4_ev4.Enabled = True
            pkm4_ev5.Enabled = True
            pkm4_ev6.Enabled = True
            pkm4_level.Enabled = True
            pkm4_helditem.Enabled = True
            pkm4_move1.Enabled = True
            pkm4_move2.Enabled = True
            pkm4_move3.Enabled = True
            pkm4_move4.Enabled = True
            pkm4_nickname.Enabled = True
            pkm4_helditem_image.Visible = True
            pkm4_shiny.Enabled = True
            pkm4_gender.Enabled = True
        Else
            pkm4_nature.Enabled = False
            pkm4_iv1.Enabled = False
            pkm4_iv2.Enabled = False
            pkm4_iv3.Enabled = False
            pkm4_iv4.Enabled = False
            pkm4_iv5.Enabled = False
            pkm4_iv6.Enabled = False
            pkm4_ev1.Enabled = False
            pkm4_ev2.Enabled = False
            pkm4_ev3.Enabled = False
            pkm4_ev4.Enabled = False
            pkm4_ev5.Enabled = False
            pkm4_ev6.Enabled = False
            pkm4_level.Enabled = False
            pkm4_helditem.Enabled = False
            pkm4_move1.Enabled = False
            pkm4_move2.Enabled = False
            pkm4_move3.Enabled = False
            pkm4_move4.Enabled = False
            pkm4_nickname.Enabled = False
            pkm4_helditem_image.Visible = False
            pkm4_shiny.Enabled = False
            pkm4_gender.Enabled = False
        End If


        For i = 1 To pkm4_weaknesses.Controls.Count
            pkm4_weaknesses.Controls(0).Dispose()
        Next
        For i = 1 To pkm4_resistances.Controls.Count
            pkm4_resistances.Controls(0).Dispose()
        Next

        Dim StartWeak As UShort = 4
        Dim StartResist As UShort = 4

        For i = 0 To 16
            If pokedata(24)(i) > 100 Then
                Dim NewWeakness As New PictureBox
                NewWeakness.Size = New System.Drawing.Size(36, 18)
                NewWeakness.Parent = pkm4_weaknesses
                NewWeakness.SizeMode = PictureBoxSizeMode.CenterImage
                NewWeakness.Image = My.Resources.ResourceManager.GetObject("type" & (i + 1))
                NewWeakness.Tag = "type" & (i + 1)
                NewWeakness.Location = New System.Drawing.Point(StartWeak, 12)
                '4x
                If pokedata(24)(i) > 200 Then
                    NewWeakness.BackColor = Color.DarkRed
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewWeakness, "4×")
                Else
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewWeakness, "2×")
                End If
                StartWeak += 38
            ElseIf pokedata(24)(i) < 100 Then
                Dim NewResistance As New PictureBox
                NewResistance.Size = New System.Drawing.Size(36, 18)
                NewResistance.Parent = pkm4_resistances
                NewResistance.SizeMode = PictureBoxSizeMode.CenterImage
                NewResistance.Image = My.Resources.ResourceManager.GetObject("type" & (i + 1))
                NewResistance.Tag = "type" & (i + 1)
                NewResistance.Location = New System.Drawing.Point(StartResist, 12)
                '1/4x or 0x
                If pokedata(24)(i) = 0 Then
                    NewResistance.BackColor = Color.Black
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "0×")
                ElseIf pokedata(24)(i) < 50 Then
                    NewResistance.BackColor = Color.Green
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "¼×")
                Else
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "½×")
                End If
                StartResist += 38
            End If
        Next

        pkm4_species.Text = pokedata(21)

        pokemonNames(3) = pokedata(5)
        UpdateTab4()

        CalcStats4()
    End Sub

    Public Sub UpdateTab4()
        If pkm4.SelectedIndex > 0 Then
            If pkm4_nickname.Text.Length > 0 Then
                TabPage4.Text = "4. " & pkm4_nickname.Text
            Else
                TabPage4.Text = "4. " & pokemonNames(3)
            End If
        Else
            TabPage4.Text = "Pokémon 4"
        End If
    End Sub

    Private Sub pkm4_iv1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm4_iv1.ValueChanged, pkm4_iv2.ValueChanged, pkm4_iv3.ValueChanged, pkm4_iv4.ValueChanged, pkm4_iv5.ValueChanged, pkm4_iv6.ValueChanged
        pkm4_ivtotal.Text = pkm4_iv1.Value + pkm4_iv2.Value + _
            pkm4_iv3.Value + pkm4_iv4.Value + _
            pkm4_iv5.Value + pkm4_iv6.Value

        CalcStats4()
    End Sub

    Private Sub pkm4_ev1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm4_ev1.ValueChanged, pkm4_ev2.ValueChanged, pkm4_ev3.ValueChanged, pkm4_ev4.ValueChanged, pkm4_ev5.ValueChanged, pkm4_ev6.ValueChanged
        Dim sum As UShort
        sum = pkm4_ev1.Value + pkm4_ev2.Value + _
            pkm4_ev3.Value + pkm4_ev4.Value + _
            pkm4_ev5.Value + pkm4_ev6.Value

        pkm4_evtotal.Text = sum
        If sum > 510 Then
            pkm4_evtotal.ForeColor = Color.Red
        Else
            pkm4_evtotal.ForeColor = Color.Black
        End If

        CalcStats4()
    End Sub

    Public Sub CalcStats4()
        Dim nature As Double = 1.0
        If Val(pkm4_hp.Text) = 0 Then
            pkm4_stat1.Text = 0
            pkm4_stat2.Text = 0
            pkm4_stat3.Text = 0
            pkm4_stat4.Text = 0
            pkm4_stat5.Text = 0
            pkm4_stat6.Text = 0

        Else
            If Val(pkm4_hp.Text) = 1 Then
                pkm4_stat1.Text = 1
            Else
                pkm4_stat1.Text = Math.Floor((pkm4_iv1.Value + (2 * Val(pkm4_hp.Text)) + pkm4_ev1.Value / 4.0 + 100.0) * pkm4_level.Value / 100.0) + 10
            End If
            pkm4_stat2.Text = Math.Floor(Math.Floor((pkm4_iv2.Value + (2 * Val(pkm4_attack.Text)) + pkm4_ev2.Value / 4.0) * pkm4_level.Value / 100.0 + 5) * pkm4_nature_multipliers(0))
            pkm4_stat3.Text = Math.Floor(Math.Floor((pkm4_iv3.Value + (2 * Val(pkm4_defense.Text)) + pkm4_ev3.Value / 4.0) * pkm4_level.Value / 100.0 + 5) * pkm4_nature_multipliers(1))
            pkm4_stat4.Text = Math.Floor(Math.Floor((pkm4_iv4.Value + (2 * Val(pkm4_specialattack.Text)) + pkm4_ev4.Value / 4.0) * pkm4_level.Value / 100.0 + 5) * pkm4_nature_multipliers(2))
            pkm4_stat5.Text = Math.Floor(Math.Floor((pkm4_iv5.Value + (2 * Val(pkm4_specialdefense.Text)) + pkm4_ev5.Value / 4.0) * pkm4_level.Value / 100.0 + 5) * pkm4_nature_multipliers(3))
            pkm4_stat6.Text = Math.Floor(Math.Floor((pkm4_iv6.Value + (2 * Val(pkm4_speed.Text)) + pkm4_ev6.Value / 4.0) * pkm4_level.Value / 100.0 + 5) * pkm4_nature_multipliers(4))

        End If

        'Set colors based on nature
        pkm4_stat2.ForeColor = Color.Black
        pkm4_stat3.ForeColor = Color.Black
        pkm4_stat4.ForeColor = Color.Black
        pkm4_stat5.ForeColor = Color.Black
        pkm4_stat6.ForeColor = Color.Black

        If pkm4_nature_multipliers(0) > 1.0 Then
            pkm4_stat2.ForeColor = Color.DarkRed
        ElseIf pkm4_nature_multipliers(1) > 1.0 Then
            pkm4_stat3.ForeColor = Color.DarkRed
        ElseIf pkm4_nature_multipliers(2) > 1.0 Then
            pkm4_stat4.ForeColor = Color.DarkRed
        ElseIf pkm4_nature_multipliers(3) > 1.0 Then
            pkm4_stat5.ForeColor = Color.DarkRed
        ElseIf pkm4_nature_multipliers(4) > 1.0 Then
            pkm4_stat6.ForeColor = Color.DarkRed
        End If
        If pkm4_nature_multipliers(0) < 1.0 Then
            pkm4_stat2.ForeColor = Color.DarkBlue
        ElseIf pkm4_nature_multipliers(1) < 1.0 Then
            pkm4_stat3.ForeColor = Color.DarkBlue
        ElseIf pkm4_nature_multipliers(2) < 1.0 Then
            pkm4_stat4.ForeColor = Color.DarkBlue
        ElseIf pkm4_nature_multipliers(3) < 1.0 Then
            pkm4_stat5.ForeColor = Color.DarkBlue
        ElseIf pkm4_nature_multipliers(4) < 1.0 Then
            pkm4_stat6.ForeColor = Color.DarkBlue
        End If


        'Calculate Hidden Power
        If Val(pkm4_hp.Text) = 0 Then
            pkm4_hiddenpower.Image = Nothing
            pkm4_hiddenpower_base.Text = ""
        Else
            Dim HiddenPower As Double = _
                (((pkm4_iv1.Value Mod 2) + 2 * (pkm4_iv2.Value Mod 2) + _
                 4 * (pkm4_iv3.Value Mod 2) + 8 * (pkm4_iv6.Value Mod 2) + _
                 16 * (pkm4_iv4.Value Mod 2) + 32 * (pkm4_iv5.Value Mod 2)) * _
                 15) / 63.0
            Dim HiddenPowerBase As Double = _
                ((Math.Floor((pkm4_iv1.Value Mod 4) / 2) + 2 * Math.Floor((pkm4_iv2.Value Mod 4) / 2) + _
                 4 * Math.Floor((pkm4_iv3.Value Mod 4) / 2) + 8 * Math.Floor((pkm4_iv6.Value Mod 4) / 2) + _
                 16 * Math.Floor((pkm4_iv4.Value Mod 4) / 2) + 32 * Math.Floor((pkm4_iv5.Value Mod 4) / 2)) * _
                 40) / 63.0 + 30

            pkm4_hiddenpower.Image = My.Resources.ResourceManager.GetObject("type" & (Math.Floor(HiddenPower) + 2))
            pkm4_hiddenpower_base.Text = Math.Floor(HiddenPowerBase)
        End If

    End Sub

    Private Sub pkm4_level_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm4_level.ValueChanged
        CalcStats4()
    End Sub

    Private Sub pkm4_nature_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm4_nature.SelectedIndexChanged
        GetNature(pkm4_nature.SelectedIndex, pkm4_nature_multipliers)
    End Sub

    Private Sub pkm4_move1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm4_move1.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm4_move1.SelectedIndex)

        pkm4_move1_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm4_move1_power.Text = "---"
        Else
            pkm4_move1_power.Text = movedata(1)
        End If
        pkm4_move1_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm4_move1_accuracy.Text = "---"
        Else
            pkm4_move1_accuracy.Text = movedata(3)
        End If
        pkm4_move1_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm4_move2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm4_move2.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm4_move2.SelectedIndex)

        pkm4_move2_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm4_move2_power.Text = "---"
        Else
            pkm4_move2_power.Text = movedata(1)
        End If
        pkm4_move2_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm4_move2_accuracy.Text = "---"
        Else
            pkm4_move2_accuracy.Text = movedata(3)
        End If
        pkm4_move2_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm4_move3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm4_move3.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm4_move3.SelectedIndex)

        pkm4_move3_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm4_move3_power.Text = "---"
        Else
            pkm4_move3_power.Text = movedata(1)
        End If
        pkm4_move3_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm4_move3_accuracy.Text = "---"
        Else
            pkm4_move3_accuracy.Text = movedata(3)
        End If
        pkm4_move3_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm4_move4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm4_move4.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm4_move4.SelectedIndex)

        pkm4_move4_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm4_move4_power.Text = "---"
        Else
            pkm4_move4_power.Text = movedata(1)
        End If
        pkm4_move4_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm4_move4_accuracy.Text = "---"
        Else
            pkm4_move4_accuracy.Text = movedata(3)
        End If
        pkm4_move4_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm4_nickname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm4_nickname.TextChanged
        UpdateTab4()
    End Sub

    Private Sub pkm4_helditem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm4_helditem.SelectedIndexChanged
        Try
            Connection.Open()
        Catch ex As Exception
            MsgBox("Could not connect to the database '" & DBFilename & "'.", MsgBoxStyle.Critical)
            Exit Sub
        End Try

        Dim Command As New System.Data.SqlServerCe.SqlCeCommand("SELECT identifier FROM items WHERE items.id = " & pkm4_helditem.SelectedIndex, Connection)
        Dim myReader As System.Data.SqlServerCe.SqlCeDataReader = Command.ExecuteReader(CommandBehavior.Default)

        pkm4_helditem_image.Image = Nothing
        While myReader.Read()
            pkm4_helditem_image.Image = My.Resources.ResourceManager.GetObject(myReader.GetString(0).Replace("-", "_"))
        End While

        myReader.Close()
        Connection.Close()
    End Sub

    'Private Sub pkm4_find_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    For Item = 0 To pkm4.Items.Count - 1
    '        If pkm4.Items(Item).ToString.ToLower.Contains(pkm4_find.Text.ToLower) Then
    '            pkm4.SelectedIndex = Item
    '            Exit Sub
    '        End If
    '    Next
    'End Sub

    Private Sub pkm5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm5.SelectedIndexChanged
        Dim pokedata() As Object = UpdateMon(pkm5.SelectedIndex)

        pkm5_selected = pokedata(0)
        pkm5_genus.Text = pokedata(1)
        pkm5_height.Text = pokedata(22)
        pkm5_weight.Text = pokedata(23)
        pkm5_image.Image = My.Resources.ResourceManager.GetObject(pokedata(4))
        pkm5_type1.Image = My.Resources.ResourceManager.GetObject("type" & pokedata(2))
        pkm5_type2.Image = My.Resources.ResourceManager.GetObject("type" & pokedata(3))

        pkm5_hp.Text = pokedata(7)
        pkm5_hp.Width = pokedata(8)
        pkm5_attack.Text = pokedata(9)
        pkm5_attack.Width = pokedata(10)
        pkm5_defense.Text = pokedata(11)
        pkm5_defense.Width = pokedata(12)
        pkm5_specialattack.Text = pokedata(13)
        pkm5_specialattack.Width = pokedata(14)
        pkm5_specialdefense.Text = pokedata(15)
        pkm5_specialdefense.Width = pokedata(16)
        pkm5_speed.Text = pokedata(17)
        pkm5_speed.Width = pokedata(18)
        pkm5_total.Text = pokedata(19)

        pkm5_ability.Items.Clear()
        For Each Item In pokedata(20)
            If Not Item Is Nothing Then pkm5_ability.Items.Add(Item)
        Next
        If pkm5_ability.Items.Count > 0 Then
            pkm5_ability.Enabled = True
            pkm5_ability.SelectedIndex = 0
        Else
            pkm5_ability.Enabled = False
        End If

        If pkm5_selected > 0 Then
            pkm5_nature.Enabled = True
            pkm5_iv1.Enabled = True
            pkm5_iv2.Enabled = True
            pkm5_iv3.Enabled = True
            pkm5_iv4.Enabled = True
            pkm5_iv5.Enabled = True
            pkm5_iv6.Enabled = True
            pkm5_ev1.Enabled = True
            pkm5_ev2.Enabled = True
            pkm5_ev3.Enabled = True
            pkm5_ev4.Enabled = True
            pkm5_ev5.Enabled = True
            pkm5_ev6.Enabled = True
            pkm5_level.Enabled = True
            pkm5_helditem.Enabled = True
            pkm5_move1.Enabled = True
            pkm5_move2.Enabled = True
            pkm5_move3.Enabled = True
            pkm5_move4.Enabled = True
            pkm5_nickname.Enabled = True
            pkm5_helditem_image.Visible = True
            pkm5_shiny.Enabled = True
            pkm5_gender.Enabled = True
        Else
            pkm5_nature.Enabled = False
            pkm5_iv1.Enabled = False
            pkm5_iv2.Enabled = False
            pkm5_iv3.Enabled = False
            pkm5_iv4.Enabled = False
            pkm5_iv5.Enabled = False
            pkm5_iv6.Enabled = False
            pkm5_ev1.Enabled = False
            pkm5_ev2.Enabled = False
            pkm5_ev3.Enabled = False
            pkm5_ev4.Enabled = False
            pkm5_ev5.Enabled = False
            pkm5_ev6.Enabled = False
            pkm5_level.Enabled = False
            pkm5_helditem.Enabled = False
            pkm5_move1.Enabled = False
            pkm5_move2.Enabled = False
            pkm5_move3.Enabled = False
            pkm5_move4.Enabled = False
            pkm5_nickname.Enabled = False
            pkm5_helditem_image.Visible = False
            pkm5_shiny.Enabled = False
            pkm5_gender.Enabled = False
        End If


        For i = 1 To pkm5_weaknesses.Controls.Count
            pkm5_weaknesses.Controls(0).Dispose()
        Next
        For i = 1 To pkm5_resistances.Controls.Count
            pkm5_resistances.Controls(0).Dispose()
        Next

        Dim StartWeak As UShort = 4
        Dim StartResist As UShort = 4

        For i = 0 To 16
            If pokedata(24)(i) > 100 Then
                Dim NewWeakness As New PictureBox
                NewWeakness.Size = New System.Drawing.Size(36, 18)
                NewWeakness.Parent = pkm5_weaknesses
                NewWeakness.SizeMode = PictureBoxSizeMode.CenterImage
                NewWeakness.Image = My.Resources.ResourceManager.GetObject("type" & (i + 1))
                NewWeakness.Tag = "type" & (i + 1)
                NewWeakness.Location = New System.Drawing.Point(StartWeak, 12)
                '4x
                If pokedata(24)(i) > 200 Then
                    NewWeakness.BackColor = Color.DarkRed
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewWeakness, "4×")
                Else
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewWeakness, "2×")
                End If
                StartWeak += 38
            ElseIf pokedata(24)(i) < 100 Then
                Dim NewResistance As New PictureBox
                NewResistance.Size = New System.Drawing.Size(36, 18)
                NewResistance.Parent = pkm5_resistances
                NewResistance.SizeMode = PictureBoxSizeMode.CenterImage
                NewResistance.Image = My.Resources.ResourceManager.GetObject("type" & (i + 1))
                NewResistance.Tag = "type" & (i + 1)
                NewResistance.Location = New System.Drawing.Point(StartResist, 12)
                '1/4x or 0x
                If pokedata(24)(i) = 0 Then
                    NewResistance.BackColor = Color.Black
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "0×")
                ElseIf pokedata(24)(i) < 50 Then
                    NewResistance.BackColor = Color.Green
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "¼×")
                Else
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "½×")
                End If
                StartResist += 38
            End If
        Next

        pkm5_species.Text = pokedata(21)

        pokemonNames(4) = pokedata(5)
        UpdateTab5()

        CalcStats5()
    End Sub

    Public Sub UpdateTab5()
        If pkm5.SelectedIndex > 0 Then
            If pkm5_nickname.Text.Length > 0 Then
                TabPage5.Text = "5. " & pkm5_nickname.Text
            Else
                TabPage5.Text = "5. " & pokemonNames(4)
            End If
        Else
            TabPage5.Text = "Pokémon 5"
        End If
    End Sub

    Private Sub pkm5_iv1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm5_iv1.ValueChanged, pkm5_iv2.ValueChanged, pkm5_iv3.ValueChanged, pkm5_iv4.ValueChanged, pkm5_iv5.ValueChanged, pkm5_iv6.ValueChanged
        pkm5_ivtotal.Text = pkm5_iv1.Value + pkm5_iv2.Value + _
            pkm5_iv3.Value + pkm5_iv4.Value + _
            pkm5_iv5.Value + pkm5_iv6.Value

        CalcStats5()
    End Sub

    Private Sub pkm5_ev1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm5_ev1.ValueChanged, pkm5_ev2.ValueChanged, pkm5_ev3.ValueChanged, pkm5_ev4.ValueChanged, pkm5_ev5.ValueChanged, pkm5_ev6.ValueChanged
        Dim sum As UShort
        sum = pkm5_ev1.Value + pkm5_ev2.Value + _
            pkm5_ev3.Value + pkm5_ev4.Value + _
            pkm5_ev5.Value + pkm5_ev6.Value

        pkm5_evtotal.Text = sum
        If sum > 510 Then
            pkm5_evtotal.ForeColor = Color.Red
        Else
            pkm5_evtotal.ForeColor = Color.Black
        End If

        CalcStats5()
    End Sub

    Public Sub CalcStats5()
        Dim nature As Double = 1.0
        If Val(pkm5_hp.Text) = 0 Then
            pkm5_stat1.Text = 0
            pkm5_stat2.Text = 0
            pkm5_stat3.Text = 0
            pkm5_stat4.Text = 0
            pkm5_stat5.Text = 0
            pkm5_stat6.Text = 0

        Else
            If Val(pkm5_hp.Text) = 1 Then
                pkm5_stat1.Text = 1
            Else
                pkm5_stat1.Text = Math.Floor((pkm5_iv1.Value + (2 * Val(pkm5_hp.Text)) + pkm5_ev1.Value / 4.0 + 100.0) * pkm5_level.Value / 100.0) + 10
            End If
            pkm5_stat2.Text = Math.Floor(Math.Floor((pkm5_iv2.Value + (2 * Val(pkm5_attack.Text)) + pkm5_ev2.Value / 4.0) * pkm5_level.Value / 100.0 + 5) * pkm5_nature_multipliers(0))
            pkm5_stat3.Text = Math.Floor(Math.Floor((pkm5_iv3.Value + (2 * Val(pkm5_defense.Text)) + pkm5_ev3.Value / 4.0) * pkm5_level.Value / 100.0 + 5) * pkm5_nature_multipliers(1))
            pkm5_stat4.Text = Math.Floor(Math.Floor((pkm5_iv4.Value + (2 * Val(pkm5_specialattack.Text)) + pkm5_ev4.Value / 4.0) * pkm5_level.Value / 100.0 + 5) * pkm5_nature_multipliers(2))
            pkm5_stat5.Text = Math.Floor(Math.Floor((pkm5_iv5.Value + (2 * Val(pkm5_specialdefense.Text)) + pkm5_ev5.Value / 4.0) * pkm5_level.Value / 100.0 + 5) * pkm5_nature_multipliers(3))
            pkm5_stat6.Text = Math.Floor(Math.Floor((pkm5_iv6.Value + (2 * Val(pkm5_speed.Text)) + pkm5_ev6.Value / 4.0) * pkm5_level.Value / 100.0 + 5) * pkm5_nature_multipliers(4))

        End If

        'Set colors based on nature
        pkm5_stat2.ForeColor = Color.Black
        pkm5_stat3.ForeColor = Color.Black
        pkm5_stat4.ForeColor = Color.Black
        pkm5_stat5.ForeColor = Color.Black
        pkm5_stat6.ForeColor = Color.Black

        If pkm5_nature_multipliers(0) > 1.0 Then
            pkm5_stat2.ForeColor = Color.DarkRed
        ElseIf pkm5_nature_multipliers(1) > 1.0 Then
            pkm5_stat3.ForeColor = Color.DarkRed
        ElseIf pkm5_nature_multipliers(2) > 1.0 Then
            pkm5_stat4.ForeColor = Color.DarkRed
        ElseIf pkm5_nature_multipliers(3) > 1.0 Then
            pkm5_stat5.ForeColor = Color.DarkRed
        ElseIf pkm5_nature_multipliers(4) > 1.0 Then
            pkm5_stat6.ForeColor = Color.DarkRed
        End If
        If pkm5_nature_multipliers(0) < 1.0 Then
            pkm5_stat2.ForeColor = Color.DarkBlue
        ElseIf pkm5_nature_multipliers(1) < 1.0 Then
            pkm5_stat3.ForeColor = Color.DarkBlue
        ElseIf pkm5_nature_multipliers(2) < 1.0 Then
            pkm5_stat4.ForeColor = Color.DarkBlue
        ElseIf pkm5_nature_multipliers(3) < 1.0 Then
            pkm5_stat5.ForeColor = Color.DarkBlue
        ElseIf pkm5_nature_multipliers(4) < 1.0 Then
            pkm5_stat6.ForeColor = Color.DarkBlue
        End If


        'Calculate Hidden Power
        If Val(pkm5_hp.Text) = 0 Then
            pkm5_hiddenpower.Image = Nothing
            pkm5_hiddenpower_base.Text = ""
        Else
            Dim HiddenPower As Double = _
                (((pkm5_iv1.Value Mod 2) + 2 * (pkm5_iv2.Value Mod 2) + _
                 4 * (pkm5_iv3.Value Mod 2) + 8 * (pkm5_iv6.Value Mod 2) + _
                 16 * (pkm5_iv4.Value Mod 2) + 32 * (pkm5_iv5.Value Mod 2)) * _
                 15) / 63.0
            Dim HiddenPowerBase As Double = _
                ((Math.Floor((pkm5_iv1.Value Mod 4) / 2) + 2 * Math.Floor((pkm5_iv2.Value Mod 4) / 2) + _
                 4 * Math.Floor((pkm5_iv3.Value Mod 4) / 2) + 8 * Math.Floor((pkm5_iv6.Value Mod 4) / 2) + _
                 16 * Math.Floor((pkm5_iv4.Value Mod 4) / 2) + 32 * Math.Floor((pkm5_iv5.Value Mod 4) / 2)) * _
                 40) / 63.0 + 30

            pkm5_hiddenpower.Image = My.Resources.ResourceManager.GetObject("type" & (Math.Floor(HiddenPower) + 2))
            pkm5_hiddenpower_base.Text = Math.Floor(HiddenPowerBase)
        End If

    End Sub

    Private Sub pkm5_level_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm5_level.ValueChanged
        CalcStats5()
    End Sub

    Private Sub pkm5_nature_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm5_nature.SelectedIndexChanged
        GetNature(pkm5_nature.SelectedIndex, pkm5_nature_multipliers)
    End Sub

    Private Sub pkm5_move1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm5_move1.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm5_move1.SelectedIndex)

        pkm5_move1_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm5_move1_power.Text = "---"
        Else
            pkm5_move1_power.Text = movedata(1)
        End If
        pkm5_move1_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm5_move1_accuracy.Text = "---"
        Else
            pkm5_move1_accuracy.Text = movedata(3)
        End If
        pkm5_move1_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm5_move2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm5_move2.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm5_move2.SelectedIndex)

        pkm5_move2_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm5_move2_power.Text = "---"
        Else
            pkm5_move2_power.Text = movedata(1)
        End If
        pkm5_move2_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm5_move2_accuracy.Text = "---"
        Else
            pkm5_move2_accuracy.Text = movedata(3)
        End If
        pkm5_move2_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm5_move3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm5_move3.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm5_move3.SelectedIndex)

        pkm5_move3_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm5_move3_power.Text = "---"
        Else
            pkm5_move3_power.Text = movedata(1)
        End If
        pkm5_move3_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm5_move3_accuracy.Text = "---"
        Else
            pkm5_move3_accuracy.Text = movedata(3)
        End If
        pkm5_move3_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm5_move4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm5_move4.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm5_move4.SelectedIndex)

        pkm5_move4_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm5_move4_power.Text = "---"
        Else
            pkm5_move4_power.Text = movedata(1)
        End If
        pkm5_move4_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm5_move4_accuracy.Text = "---"
        Else
            pkm5_move4_accuracy.Text = movedata(3)
        End If
        pkm5_move4_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm5_nickname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm5_nickname.TextChanged
        UpdateTab5()
    End Sub

    Private Sub pkm5_helditem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm5_helditem.SelectedIndexChanged
        Try
            Connection.Open()
        Catch ex As Exception
            MsgBox("Could not connect to the database '" & DBFilename & "'.", MsgBoxStyle.Critical)
            Exit Sub
        End Try

        Dim Command As New System.Data.SqlServerCe.SqlCeCommand("SELECT identifier FROM items WHERE items.id = " & pkm5_helditem.SelectedIndex, Connection)
        Dim myReader As System.Data.SqlServerCe.SqlCeDataReader = Command.ExecuteReader(CommandBehavior.Default)

        pkm5_helditem_image.Image = Nothing
        While myReader.Read()
            pkm5_helditem_image.Image = My.Resources.ResourceManager.GetObject(myReader.GetString(0).Replace("-", "_"))
        End While

        myReader.Close()
        Connection.Close()
    End Sub

    'Private Sub pkm5_find_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    For Item = 0 To pkm5.Items.Count - 1
    '        If pkm5.Items(Item).ToString.ToLower.Contains(pkm5_find.Text.ToLower) Then
    '            pkm5.SelectedIndex = Item
    '            Exit Sub
    '        End If
    '    Next
    'End Sub

    Private Sub pkm6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm6.SelectedIndexChanged
        Dim pokedata() As Object = UpdateMon(pkm6.SelectedIndex)

        pkm6_selected = pokedata(0)
        pkm6_genus.Text = pokedata(1)
        pkm6_height.Text = pokedata(22)
        pkm6_weight.Text = pokedata(23)
        pkm6_image.Image = My.Resources.ResourceManager.GetObject(pokedata(4))
        pkm6_type1.Image = My.Resources.ResourceManager.GetObject("type" & pokedata(2))
        pkm6_type2.Image = My.Resources.ResourceManager.GetObject("type" & pokedata(3))

        pkm6_hp.Text = pokedata(7)
        pkm6_hp.Width = pokedata(8)
        pkm6_attack.Text = pokedata(9)
        pkm6_attack.Width = pokedata(10)
        pkm6_defense.Text = pokedata(11)
        pkm6_defense.Width = pokedata(12)
        pkm6_specialattack.Text = pokedata(13)
        pkm6_specialattack.Width = pokedata(14)
        pkm6_specialdefense.Text = pokedata(15)
        pkm6_specialdefense.Width = pokedata(16)
        pkm6_speed.Text = pokedata(17)
        pkm6_speed.Width = pokedata(18)
        pkm6_total.Text = pokedata(19)

        pkm6_ability.Items.Clear()
        For Each Item In pokedata(20)
            If Not Item Is Nothing Then pkm6_ability.Items.Add(Item)
        Next
        If pkm6_ability.Items.Count > 0 Then
            pkm6_ability.Enabled = True
            pkm6_ability.SelectedIndex = 0
        Else
            pkm6_ability.Enabled = False
        End If

        If pkm6_selected > 0 Then
            pkm6_nature.Enabled = True
            pkm6_iv1.Enabled = True
            pkm6_iv2.Enabled = True
            pkm6_iv3.Enabled = True
            pkm6_iv4.Enabled = True
            pkm6_iv5.Enabled = True
            pkm6_iv6.Enabled = True
            pkm6_ev1.Enabled = True
            pkm6_ev2.Enabled = True
            pkm6_ev3.Enabled = True
            pkm6_ev4.Enabled = True
            pkm6_ev5.Enabled = True
            pkm6_ev6.Enabled = True
            pkm6_level.Enabled = True
            pkm6_helditem.Enabled = True
            pkm6_move1.Enabled = True
            pkm6_move2.Enabled = True
            pkm6_move3.Enabled = True
            pkm6_move4.Enabled = True
            pkm6_nickname.Enabled = True
            pkm6_helditem_image.Visible = True
            pkm6_shiny.Enabled = True
            pkm6_gender.Enabled = True
        Else
            pkm6_nature.Enabled = False
            pkm6_iv1.Enabled = False
            pkm6_iv2.Enabled = False
            pkm6_iv3.Enabled = False
            pkm6_iv4.Enabled = False
            pkm6_iv5.Enabled = False
            pkm6_iv6.Enabled = False
            pkm6_ev1.Enabled = False
            pkm6_ev2.Enabled = False
            pkm6_ev3.Enabled = False
            pkm6_ev4.Enabled = False
            pkm6_ev5.Enabled = False
            pkm6_ev6.Enabled = False
            pkm6_level.Enabled = False
            pkm6_helditem.Enabled = False
            pkm6_move1.Enabled = False
            pkm6_move2.Enabled = False
            pkm6_move3.Enabled = False
            pkm6_move4.Enabled = False
            pkm6_nickname.Enabled = False
            pkm6_helditem_image.Visible = False
            pkm6_shiny.Enabled = False
            pkm6_gender.Enabled = False
        End If


        For i = 1 To pkm6_weaknesses.Controls.Count
            pkm6_weaknesses.Controls(0).Dispose()
        Next
        For i = 1 To pkm6_resistances.Controls.Count
            pkm6_resistances.Controls(0).Dispose()
        Next

        Dim StartWeak As UShort = 4
        Dim StartResist As UShort = 4

        For i = 0 To 16
            If pokedata(24)(i) > 100 Then
                Dim NewWeakness As New PictureBox
                NewWeakness.Size = New System.Drawing.Size(36, 18)
                NewWeakness.Parent = pkm6_weaknesses
                NewWeakness.SizeMode = PictureBoxSizeMode.CenterImage
                NewWeakness.Image = My.Resources.ResourceManager.GetObject("type" & (i + 1))
                NewWeakness.Tag = "type" & (i + 1)
                NewWeakness.Location = New System.Drawing.Point(StartWeak, 12)
                '4x
                If pokedata(24)(i) > 200 Then
                    NewWeakness.BackColor = Color.DarkRed
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewWeakness, "4×")
                Else
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewWeakness, "2×")
                End If
                StartWeak += 38
            ElseIf pokedata(24)(i) < 100 Then
                Dim NewResistance As New PictureBox
                NewResistance.Size = New System.Drawing.Size(36, 18)
                NewResistance.Parent = pkm6_resistances
                NewResistance.SizeMode = PictureBoxSizeMode.CenterImage
                NewResistance.Image = My.Resources.ResourceManager.GetObject("type" & (i + 1))
                NewResistance.Tag = "type" & (i + 1)
                NewResistance.Location = New System.Drawing.Point(StartResist, 12)
                '1/4x or 0x
                If pokedata(24)(i) = 0 Then
                    NewResistance.BackColor = Color.Black
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "0×")
                ElseIf pokedata(24)(i) < 50 Then
                    NewResistance.BackColor = Color.Green
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "¼×")
                Else
                    Dim NewTip As New ToolTip
                    NewTip.SetToolTip(NewResistance, "½×")
                End If
                StartResist += 38
            End If
        Next

        pkm6_species.Text = pokedata(21)

        pokemonNames(5) = pokedata(5)
        UpdateTab6()

        CalcStats6()
    End Sub

    Public Sub UpdateTab6()
        If pkm6.SelectedIndex > 0 Then
            If pkm6_nickname.Text.Length > 0 Then
                TabPage6.Text = "6. " & pkm6_nickname.Text
            Else
                TabPage6.Text = "6. " & pokemonNames(5)
            End If
        Else
            TabPage6.Text = "Pokémon 6"
        End If
    End Sub

    Private Sub pkm6_iv1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm6_iv1.ValueChanged, pkm6_iv2.ValueChanged, pkm6_iv3.ValueChanged, pkm6_iv4.ValueChanged, pkm6_iv5.ValueChanged, pkm6_iv6.ValueChanged
        pkm6_ivtotal.Text = pkm6_iv1.Value + pkm6_iv2.Value + _
            pkm6_iv3.Value + pkm6_iv4.Value + _
            pkm6_iv5.Value + pkm6_iv6.Value

        CalcStats6()
    End Sub

    Private Sub pkm6_ev1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm6_ev1.ValueChanged, pkm6_ev2.ValueChanged, pkm6_ev3.ValueChanged, pkm6_ev4.ValueChanged, pkm6_ev5.ValueChanged, pkm6_ev6.ValueChanged
        Dim sum As UShort
        sum = pkm6_ev1.Value + pkm6_ev2.Value + _
            pkm6_ev3.Value + pkm6_ev4.Value + _
            pkm6_ev5.Value + pkm6_ev6.Value

        pkm6_evtotal.Text = sum
        If sum > 510 Then
            pkm6_evtotal.ForeColor = Color.Red
        Else
            pkm6_evtotal.ForeColor = Color.Black
        End If

        CalcStats6()
    End Sub

    Public Sub CalcStats6()
        Dim nature As Double = 1.0
        If Val(pkm6_hp.Text) = 0 Then
            pkm6_stat1.Text = 0
            pkm6_stat2.Text = 0
            pkm6_stat3.Text = 0
            pkm6_stat4.Text = 0
            pkm6_stat5.Text = 0
            pkm6_stat6.Text = 0

        Else
            If Val(pkm6_hp.Text) = 1 Then
                pkm6_stat1.Text = 1
            Else
                pkm6_stat1.Text = Math.Floor((pkm6_iv1.Value + (2 * Val(pkm6_hp.Text)) + pkm6_ev1.Value / 4.0 + 100.0) * pkm6_level.Value / 100.0) + 10
            End If
            pkm6_stat2.Text = Math.Floor(Math.Floor((pkm6_iv2.Value + (2 * Val(pkm6_attack.Text)) + pkm6_ev2.Value / 4.0) * pkm6_level.Value / 100.0 + 5) * pkm6_nature_multipliers(0))
            pkm6_stat3.Text = Math.Floor(Math.Floor((pkm6_iv3.Value + (2 * Val(pkm6_defense.Text)) + pkm6_ev3.Value / 4.0) * pkm6_level.Value / 100.0 + 5) * pkm6_nature_multipliers(1))
            pkm6_stat4.Text = Math.Floor(Math.Floor((pkm6_iv4.Value + (2 * Val(pkm6_specialattack.Text)) + pkm6_ev4.Value / 4.0) * pkm6_level.Value / 100.0 + 5) * pkm6_nature_multipliers(2))
            pkm6_stat5.Text = Math.Floor(Math.Floor((pkm6_iv5.Value + (2 * Val(pkm6_specialdefense.Text)) + pkm6_ev5.Value / 4.0) * pkm6_level.Value / 100.0 + 5) * pkm6_nature_multipliers(3))
            pkm6_stat6.Text = Math.Floor(Math.Floor((pkm6_iv6.Value + (2 * Val(pkm6_speed.Text)) + pkm6_ev6.Value / 4.0) * pkm6_level.Value / 100.0 + 5) * pkm6_nature_multipliers(4))

        End If

        'Set colors based on nature
        pkm6_stat2.ForeColor = Color.Black
        pkm6_stat3.ForeColor = Color.Black
        pkm6_stat4.ForeColor = Color.Black
        pkm6_stat5.ForeColor = Color.Black
        pkm6_stat6.ForeColor = Color.Black

        If pkm6_nature_multipliers(0) > 1.0 Then
            pkm6_stat2.ForeColor = Color.DarkRed
        ElseIf pkm6_nature_multipliers(1) > 1.0 Then
            pkm6_stat3.ForeColor = Color.DarkRed
        ElseIf pkm6_nature_multipliers(2) > 1.0 Then
            pkm6_stat4.ForeColor = Color.DarkRed
        ElseIf pkm6_nature_multipliers(3) > 1.0 Then
            pkm6_stat5.ForeColor = Color.DarkRed
        ElseIf pkm6_nature_multipliers(4) > 1.0 Then
            pkm6_stat6.ForeColor = Color.DarkRed
        End If
        If pkm6_nature_multipliers(0) < 1.0 Then
            pkm6_stat2.ForeColor = Color.DarkBlue
        ElseIf pkm6_nature_multipliers(1) < 1.0 Then
            pkm6_stat3.ForeColor = Color.DarkBlue
        ElseIf pkm6_nature_multipliers(2) < 1.0 Then
            pkm6_stat4.ForeColor = Color.DarkBlue
        ElseIf pkm6_nature_multipliers(3) < 1.0 Then
            pkm6_stat5.ForeColor = Color.DarkBlue
        ElseIf pkm6_nature_multipliers(4) < 1.0 Then
            pkm6_stat6.ForeColor = Color.DarkBlue
        End If


        'Calculate Hidden Power
        If Val(pkm6_hp.Text) = 0 Then
            pkm6_hiddenpower.Image = Nothing
            pkm6_hiddenpower_base.Text = ""
        Else
            Dim HiddenPower As Double = _
                (((pkm6_iv1.Value Mod 2) + 2 * (pkm6_iv2.Value Mod 2) + _
                 4 * (pkm6_iv3.Value Mod 2) + 8 * (pkm6_iv6.Value Mod 2) + _
                 16 * (pkm6_iv4.Value Mod 2) + 32 * (pkm6_iv5.Value Mod 2)) * _
                 15) / 63.0
            Dim HiddenPowerBase As Double = _
                ((Math.Floor((pkm6_iv1.Value Mod 4) / 2) + 2 * Math.Floor((pkm6_iv2.Value Mod 4) / 2) + _
                 4 * Math.Floor((pkm6_iv3.Value Mod 4) / 2) + 8 * Math.Floor((pkm6_iv6.Value Mod 4) / 2) + _
                 16 * Math.Floor((pkm6_iv4.Value Mod 4) / 2) + 32 * Math.Floor((pkm6_iv5.Value Mod 4) / 2)) * _
                 40) / 63.0 + 30

            pkm6_hiddenpower.Image = My.Resources.ResourceManager.GetObject("type" & (Math.Floor(HiddenPower) + 2))
            pkm6_hiddenpower_base.Text = Math.Floor(HiddenPowerBase)
        End If

    End Sub

    Private Sub pkm6_level_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm6_level.ValueChanged
        CalcStats6()
    End Sub

    Private Sub pkm6_nature_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm6_nature.SelectedIndexChanged
        GetNature(pkm6_nature.SelectedIndex, pkm6_nature_multipliers)
    End Sub

    Private Sub pkm6_move1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm6_move1.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm6_move1.SelectedIndex)

        pkm6_move1_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm6_move1_power.Text = "---"
        Else
            pkm6_move1_power.Text = movedata(1)
        End If
        pkm6_move1_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm6_move1_accuracy.Text = "---"
        Else
            pkm6_move1_accuracy.Text = movedata(3)
        End If
        pkm6_move1_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm6_move2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm6_move2.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm6_move2.SelectedIndex)

        pkm6_move2_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm6_move2_power.Text = "---"
        Else
            pkm6_move2_power.Text = movedata(1)
        End If
        pkm6_move2_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm6_move2_accuracy.Text = "---"
        Else
            pkm6_move2_accuracy.Text = movedata(3)
        End If
        pkm6_move2_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm6_move3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm6_move3.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm6_move3.SelectedIndex)

        pkm6_move3_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm6_move3_power.Text = "---"
        Else
            pkm6_move3_power.Text = movedata(1)
        End If
        pkm6_move3_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm6_move3_accuracy.Text = "---"
        Else
            pkm6_move3_accuracy.Text = movedata(3)
        End If
        pkm6_move3_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm6_move4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm6_move4.SelectedIndexChanged
        Dim movedata() As Byte = GetMove(pkm6_move4.SelectedIndex)

        pkm6_move4_type.Image = My.Resources.ResourceManager.GetObject("type" & movedata(0))
        If movedata(1) = 0 Then
            pkm6_move4_power.Text = "---"
        Else
            pkm6_move4_power.Text = movedata(1)
        End If
        pkm6_move4_pp.Text = movedata(2)
        If movedata(3) = 0 Then
            pkm6_move4_accuracy.Text = "---"
        Else
            pkm6_move4_accuracy.Text = movedata(3)
        End If
        pkm6_move4_cat.Image = My.Resources.ResourceManager.GetObject("cat" & movedata(4))
    End Sub

    Private Sub pkm6_nickname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm6_nickname.TextChanged
        UpdateTab6()
    End Sub

    Private Sub pkm6_helditem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pkm6_helditem.SelectedIndexChanged
        Try
            Connection.Open()
        Catch ex As Exception
            MsgBox("Could not connect to the database '" & DBFilename & "'.", MsgBoxStyle.Critical)
            Exit Sub
        End Try

        Dim Command As New System.Data.SqlServerCe.SqlCeCommand("SELECT identifier FROM items WHERE items.id = " & pkm6_helditem.SelectedIndex, Connection)
        Dim myReader As System.Data.SqlServerCe.SqlCeDataReader = Command.ExecuteReader(CommandBehavior.Default)

        pkm6_helditem_image.Image = Nothing
        While myReader.Read()
            pkm6_helditem_image.Image = My.Resources.ResourceManager.GetObject(myReader.GetString(0).Replace("-", "_"))
        End While

        myReader.Close()
        Connection.Close()
    End Sub

    'Private Sub pkm6_find_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    For Item = 0 To pkm6.Items.Count - 1
    '        If pkm6.Items(Item).ToString.ToLower.Contains(pkm6_find.Text.ToLower) Then
    '            pkm6.SelectedIndex = Item
    '            Exit Sub
    '        End If
    '    Next
    'End Sub

    Private Sub TabPage7_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage7.Enter
        Analyze()
    End Sub

    Public Sub Analyze()
        Cursor = Cursors.WaitCursor
        LoadPokémonToolStripMenuItem.Enabled = False
        SavePokToolStripMenuItem.Enabled = False
        DeletePokémonToolStripMenuItem.Enabled = False
        Dim valid = True
        For i = 0 To 4
            For j = i + 1 To 5
                If pokemonNames(i) <> "" And pokemonNames(j) <> "" And pokemonNames(i) = pokemonNames(j) Then
                    valid = False
                End If
            Next
        Next
        If Not valid Then
            species_clause.Text = "Fail"
            species_clause.ForeColor = Color.Red
        Else
            species_clause.Text = "Pass"
            species_clause.ForeColor = Color.Green
        End If

        If (pkm1_helditem.SelectedIndex <> 0 And pkm1.SelectedIndex > 0 And _
            pkm2_helditem.SelectedIndex <> 0 And pkm2.SelectedIndex > 0 And _
            pkm1_helditem.SelectedIndex = pkm2_helditem.SelectedIndex) Or _
            (pkm1_helditem.SelectedIndex <> 0 And pkm1.SelectedIndex > 0 And _
            pkm3_helditem.SelectedIndex <> 0 And pkm3.SelectedIndex > 0 And _
            pkm1_helditem.SelectedIndex = pkm3_helditem.SelectedIndex) Or _
            (pkm1_helditem.SelectedIndex <> 0 And pkm1.SelectedIndex > 0 And _
            pkm4_helditem.SelectedIndex <> 0 And pkm4.SelectedIndex > 0 And _
            pkm1_helditem.SelectedIndex = pkm4_helditem.SelectedIndex) Or _
            (pkm1_helditem.SelectedIndex <> 0 And pkm1.SelectedIndex > 0 And _
            pkm5_helditem.SelectedIndex <> 0 And pkm5.SelectedIndex > 0 And _
            pkm1_helditem.SelectedIndex = pkm5_helditem.SelectedIndex) Or _
            (pkm1_helditem.SelectedIndex <> 0 And pkm1.SelectedIndex > 0 And _
            pkm6_helditem.SelectedIndex <> 0 And pkm6.SelectedIndex > 0 And _
            pkm1_helditem.SelectedIndex = pkm6_helditem.SelectedIndex) Or _
            (pkm2_helditem.SelectedIndex <> 0 And pkm2.SelectedIndex > 0 And _
            pkm3_helditem.SelectedIndex <> 0 And pkm3.SelectedIndex > 0 And _
            pkm2_helditem.SelectedIndex = pkm3_helditem.SelectedIndex) Or _
            (pkm2_helditem.SelectedIndex <> 0 And pkm2.SelectedIndex > 0 And _
            pkm4_helditem.SelectedIndex <> 0 And pkm4.SelectedIndex > 0 And _
            pkm2_helditem.SelectedIndex = pkm4_helditem.SelectedIndex) Or _
            (pkm2_helditem.SelectedIndex <> 0 And pkm2.SelectedIndex > 0 And _
            pkm5_helditem.SelectedIndex <> 0 And pkm5.SelectedIndex > 0 And _
            pkm2_helditem.SelectedIndex = pkm5_helditem.SelectedIndex) Or _
            (pkm2_helditem.SelectedIndex <> 0 And pkm2.SelectedIndex > 0 And _
            pkm6_helditem.SelectedIndex <> 0 And pkm6.SelectedIndex > 0 And _
            pkm2_helditem.SelectedIndex = pkm6_helditem.SelectedIndex) Or _
            (pkm3_helditem.SelectedIndex <> 0 And pkm3.SelectedIndex > 0 And _
            pkm4_helditem.SelectedIndex <> 0 And pkm4.SelectedIndex > 0 And _
            pkm3_helditem.SelectedIndex = pkm4_helditem.SelectedIndex) Or _
            (pkm3_helditem.SelectedIndex <> 0 And pkm3.SelectedIndex > 0 And _
            pkm5_helditem.SelectedIndex <> 0 And pkm5.SelectedIndex > 0 And _
            pkm3_helditem.SelectedIndex = pkm5_helditem.SelectedIndex) Or _
            (pkm3_helditem.SelectedIndex <> 0 And pkm3.SelectedIndex > 0 And _
            pkm6_helditem.SelectedIndex <> 0 And pkm6.SelectedIndex > 0 And _
            pkm3_helditem.SelectedIndex = pkm6_helditem.SelectedIndex) Or _
            (pkm4_helditem.SelectedIndex <> 0 And pkm4.SelectedIndex > 0 And _
            pkm5_helditem.SelectedIndex <> 0 And pkm5.SelectedIndex > 0 And _
            pkm4_helditem.SelectedIndex = pkm5_helditem.SelectedIndex) Or _
            (pkm4_helditem.SelectedIndex <> 0 And pkm4.SelectedIndex > 0 And _
            pkm6_helditem.SelectedIndex <> 0 And pkm6.SelectedIndex > 0 And _
            pkm4_helditem.SelectedIndex = pkm6_helditem.SelectedIndex) Or _
            (pkm5_helditem.SelectedIndex <> 0 And pkm5.SelectedIndex > 0 And _
            pkm6_helditem.SelectedIndex <> 0 And pkm6.SelectedIndex > 0 And _
            pkm5_helditem.SelectedIndex = pkm6_helditem.SelectedIndex) Then

            item_clause.Text = "Fail"
            item_clause.ForeColor = Color.Red
        Else
            item_clause.Text = "Pass"
            item_clause.ForeColor = Color.Green
        End If

        'Clear labels in grid
        For Each Ctrl In TableLayoutPanel1.Controls
            If Ctrl.Name.Contains("type") Then
                Ctrl.Text = ""
                Ctrl.forecolor = Color.Black
                Ctrl.BackColor = Color.Transparent
                Ctrl.Font = New Font(prototype:=Ctrl.Font, newStyle:=FontStyle.Regular)
            End If
        Next

        For Each WkCtrl In {pkm1_weaknesses, pkm2_weaknesses, _
                          pkm3_weaknesses, pkm4_weaknesses, _
                          pkm5_weaknesses, pkm6_weaknesses}
            For Each Ctrl In WkCtrl.Controls
                Select Case Ctrl.BackColor
                    Case Color.DarkSlateBlue
                        TableLayoutPanel1.Controls(Ctrl.Tag & "_4").Text = Val(TableLayoutPanel1.Controls(Ctrl.Tag & "_4").Text) + 1
                   Case Else
                        TableLayoutPanel1.Controls(Ctrl.Tag & "_2").Text = Val(TableLayoutPanel1.Controls(Ctrl.Tag & "_2").Text) + 1
                End Select
                If Val(TableLayoutPanel1.Controls(Ctrl.Tag & "_4").Text) + Val(TableLayoutPanel1.Controls(Ctrl.Tag & "_2").Text) >= 3 Then
                    TableLayoutPanel1.Controls(Ctrl.Tag & "_2").BackColor = Color.LightPink
                    TableLayoutPanel1.Controls(Ctrl.Tag & "_2").ForeColor = Color.Red
                    TableLayoutPanel1.Controls(Ctrl.Tag & "_2").Font = New Font(prototype:=TableLayoutPanel1.Controls(Ctrl.Tag & "_2").Font, newStyle:=FontStyle.Bold)
                    TableLayoutPanel1.Controls(Ctrl.Tag & "_4").BackColor = Color.LightPink
                    TableLayoutPanel1.Controls(Ctrl.Tag & "_4").ForeColor = Color.Red
                    TableLayoutPanel1.Controls(Ctrl.Tag & "_4").Font = New Font(prototype:=TableLayoutPanel1.Controls(Ctrl.Tag & "_4").Font, newStyle:=FontStyle.Bold)
                End If
            Next Ctrl
        Next WkCtrl
        For Each RsCtrl In {pkm1_resistances, pkm2_resistances, _
                          pkm3_resistances, pkm4_resistances, _
                          pkm5_resistances, pkm6_resistances}
            For Each Ctrl In RsCtrl.Controls
                Select Case Ctrl.BackColor
                    Case Color.DarkRed
                        TableLayoutPanel1.Controls(Ctrl.Tag & "_14").Text = Val(TableLayoutPanel1.Controls(Ctrl.Tag & "_14").Text) + 1
                    Case Color.Black
                        TableLayoutPanel1.Controls(Ctrl.Tag & "_0").Text = Val(TableLayoutPanel1.Controls(Ctrl.Tag & "_0").Text) + 1
                     Case Else
                        TableLayoutPanel1.Controls(Ctrl.Tag & "_12").Text = Val(TableLayoutPanel1.Controls(Ctrl.Tag & "_12").Text) + 1
                End Select
                If Val(TableLayoutPanel1.Controls(Ctrl.Tag & "_14").Text) + Val(TableLayoutPanel1.Controls(Ctrl.Tag & "_12").Text) + Val(TableLayoutPanel1.Controls(Ctrl.Tag & "_0").Text) >= 3 Then
                    TableLayoutPanel1.Controls(Ctrl.Tag & "_14").BackColor = Color.LightGreen
                    TableLayoutPanel1.Controls(Ctrl.Tag & "_14").Font = New Font(prototype:=TableLayoutPanel1.Controls(Ctrl.Tag & "_14").Font, newStyle:=FontStyle.Bold)
                    TableLayoutPanel1.Controls(Ctrl.Tag & "_12").BackColor = Color.LightGreen
                    TableLayoutPanel1.Controls(Ctrl.Tag & "_12").Font = New Font(prototype:=TableLayoutPanel1.Controls(Ctrl.Tag & "_12").Font, newStyle:=FontStyle.Bold)
                    TableLayoutPanel1.Controls(Ctrl.Tag & "_0").BackColor = Color.LightGreen
                    TableLayoutPanel1.Controls(Ctrl.Tag & "_0").Font = New Font(prototype:=TableLayoutPanel1.Controls(Ctrl.Tag & "_0").Font, newStyle:=FontStyle.Bold)
                End If
            Next Ctrl
        Next RsCtrl

        Dim ActivePokemon As Byte = 0
        Dim StatSum(5) As UShort
        Dim BaseSum(5) As UShort
        If pkm1.SelectedIndex > 0 Then
            BaseSum(0) += Val(pkm1_hp.Text)
            BaseSum(1) += Val(pkm1_attack.Text)
            BaseSum(2) += Val(pkm1_defense.Text)
            BaseSum(3) += Val(pkm1_specialattack.Text)
            BaseSum(4) += Val(pkm1_specialdefense.Text)
            BaseSum(5) += Val(pkm1_speed.Text)
            StatSum(0) += Val(pkm1_stat1.Text)
            StatSum(1) += Val(pkm1_stat2.Text)
            StatSum(2) += Val(pkm1_stat3.Text)
            StatSum(3) += Val(pkm1_stat4.Text)
            StatSum(4) += Val(pkm1_stat5.Text)
            StatSum(5) += Val(pkm1_stat6.Text)
            ActivePokemon += 1
        End If
        If pkm2.SelectedIndex > 0 Then
            BaseSum(0) += Val(pkm2_hp.Text)
            BaseSum(1) += Val(pkm2_attack.Text)
            BaseSum(2) += Val(pkm2_defense.Text)
            BaseSum(3) += Val(pkm2_specialattack.Text)
            BaseSum(4) += Val(pkm2_specialdefense.Text)
            BaseSum(5) += Val(pkm2_speed.Text)
            StatSum(0) += Val(pkm2_stat1.Text)
            StatSum(1) += Val(pkm2_stat2.Text)
            StatSum(2) += Val(pkm2_stat3.Text)
            StatSum(3) += Val(pkm2_stat4.Text)
            StatSum(4) += Val(pkm2_stat5.Text)
            StatSum(5) += Val(pkm2_stat6.Text)
            ActivePokemon += 1
        End If
        If pkm3.SelectedIndex > 0 Then
            BaseSum(0) += Val(pkm3_hp.Text)
            BaseSum(1) += Val(pkm3_attack.Text)
            BaseSum(2) += Val(pkm3_defense.Text)
            BaseSum(3) += Val(pkm3_specialattack.Text)
            BaseSum(4) += Val(pkm3_specialdefense.Text)
            BaseSum(5) += Val(pkm3_speed.Text)
            StatSum(0) += Val(pkm3_stat1.Text)
            StatSum(1) += Val(pkm3_stat2.Text)
            StatSum(2) += Val(pkm3_stat3.Text)
            StatSum(3) += Val(pkm3_stat4.Text)
            StatSum(4) += Val(pkm3_stat5.Text)
            StatSum(5) += Val(pkm3_stat6.Text)
            ActivePokemon += 1
        End If
        If pkm4.SelectedIndex > 0 Then
            BaseSum(0) += Val(pkm4_hp.Text)
            BaseSum(1) += Val(pkm4_attack.Text)
            BaseSum(2) += Val(pkm4_defense.Text)
            BaseSum(3) += Val(pkm4_specialattack.Text)
            BaseSum(4) += Val(pkm4_specialdefense.Text)
            BaseSum(5) += Val(pkm4_speed.Text)
            StatSum(0) += Val(pkm4_stat1.Text)
            StatSum(1) += Val(pkm4_stat2.Text)
            StatSum(2) += Val(pkm4_stat3.Text)
            StatSum(3) += Val(pkm4_stat4.Text)
            StatSum(4) += Val(pkm4_stat5.Text)
            StatSum(5) += Val(pkm4_stat6.Text)
            ActivePokemon += 1
        End If
        If pkm5.SelectedIndex > 0 Then
            BaseSum(0) += Val(pkm5_hp.Text)
            BaseSum(1) += Val(pkm5_attack.Text)
            BaseSum(2) += Val(pkm5_defense.Text)
            BaseSum(3) += Val(pkm5_specialattack.Text)
            BaseSum(4) += Val(pkm5_specialdefense.Text)
            BaseSum(5) += Val(pkm5_speed.Text)
            StatSum(0) += Val(pkm5_stat1.Text)
            StatSum(1) += Val(pkm5_stat2.Text)
            StatSum(2) += Val(pkm5_stat3.Text)
            StatSum(3) += Val(pkm5_stat4.Text)
            StatSum(4) += Val(pkm5_stat5.Text)
            StatSum(5) += Val(pkm5_stat6.Text)
            ActivePokemon += 1
        End If
        If pkm6.SelectedIndex > 0 Then
            BaseSum(0) += Val(pkm6_hp.Text)
            BaseSum(1) += Val(pkm6_attack.Text)
            BaseSum(2) += Val(pkm6_defense.Text)
            BaseSum(3) += Val(pkm6_specialattack.Text)
            BaseSum(4) += Val(pkm6_specialdefense.Text)
            BaseSum(5) += Val(pkm6_speed.Text)
            StatSum(0) += Val(pkm6_stat1.Text)
            StatSum(1) += Val(pkm6_stat2.Text)
            StatSum(2) += Val(pkm6_stat3.Text)
            StatSum(3) += Val(pkm6_stat4.Text)
            StatSum(4) += Val(pkm6_stat5.Text)
            StatSum(5) += Val(pkm6_stat6.Text)
            ActivePokemon += 1
        End If
        If ActivePokemon > 0 Then
            total_hp.Text = Math.Round(BaseSum(0) / ActivePokemon, 0)
            total_attack.Text = Math.Round(BaseSum(1) / ActivePokemon)
            total_defense.Text = Math.Round(BaseSum(2) / ActivePokemon)
            total_specialattack.Text = Math.Round(BaseSum(3) / ActivePokemon)
            total_specialdefense.Text = Math.Round(BaseSum(4) / ActivePokemon)
            total_speed.Text = Math.Round(BaseSum(5) / ActivePokemon)
            stat_hp.Text = Math.Round(StatSum(0) / ActivePokemon)
            stat_attack.Text = Math.Round(StatSum(1) / ActivePokemon)
            stat_defense.Text = Math.Round(StatSum(2) / ActivePokemon)
            stat_specialattack.Text = Math.Round(StatSum(3) / ActivePokemon)
            stat_specialdefense.Text = Math.Round(StatSum(4) / ActivePokemon)
            stat_speed.Text = Math.Round(StatSum(5) / ActivePokemon)
            total_stats.Text = Math.Round(BaseSum(0) / ActivePokemon, 0) + _
                Math.Round(BaseSum(1) / ActivePokemon, 0) + Math.Round(BaseSum(2) / ActivePokemon, 0) + _
                Math.Round(BaseSum(3) / ActivePokemon, 0) + Math.Round(BaseSum(4) / ActivePokemon, 0) + _
                Math.Round(BaseSum(5) / ActivePokemon, 0)
            stat_total.Text = Math.Round(StatSum(0) / ActivePokemon) + _
                Math.Round(StatSum(1) / ActivePokemon) + Math.Round(StatSum(2) / ActivePokemon) + _
                Math.Round(StatSum(3) / ActivePokemon) + Math.Round(StatSum(4) / ActivePokemon) + _
                Math.Round(StatSum(5) / ActivePokemon)
        Else
            total_hp.Text = ""
            total_attack.Text = ""
            total_defense.Text = ""
            total_specialattack.Text = ""
            total_specialdefense.Text = ""
            total_speed.Text = ""
            stat_hp.Text = ""
            stat_attack.Text = ""
            stat_defense.Text = ""
            stat_specialattack.Text = ""
            stat_specialdefense.Text = ""
            stat_speed.Text = ""
            total_stats.Text = ""
            stat_total.Text = ""
        End If

        total_pokemon.Text = ActivePokemon

        Cursor = Cursors.Default
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        If filename = "" Then
            If GetFilename() Then
                DoSave()
            End If
        Else
            DoSave()
        End If
    End Sub
    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        If GetFilename() Then
            DoSave()
        End If
    End Sub

    Public Function GetFilename()
        If SaveTeamDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            filename = SaveTeamDialog.FileName
            prettyFilename = filename.Substring(filename.LastIndexOf("\") + 1, filename.Length - filename.LastIndexOf("\") - 1)
            If prettyFilename.EndsWith(".tbt") Then prettyFilename = prettyFilename.Substring(0, prettyFilename.Length - 4)
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub DoSave()
        Dim byt(11)() As Byte
        '0-5 = notes
        byt(0) = System.Text.Encoding.UTF8.GetBytes(pkm1_notes.Text)
        byt(1) = System.Text.Encoding.UTF8.GetBytes(pkm2_notes.Text)
        byt(2) = System.Text.Encoding.UTF8.GetBytes(pkm3_notes.Text)
        byt(3) = System.Text.Encoding.UTF8.GetBytes(pkm4_notes.Text)
        byt(4) = System.Text.Encoding.UTF8.GetBytes(pkm5_notes.Text)
        byt(5) = System.Text.Encoding.UTF8.GetBytes(pkm6_notes.Text)
        '6-11 = nicknames
        byt(6) = System.Text.Encoding.UTF8.GetBytes(pkm1_nickname.Text)
        byt(7) = System.Text.Encoding.UTF8.GetBytes(pkm2_nickname.Text)
        byt(8) = System.Text.Encoding.UTF8.GetBytes(pkm3_nickname.Text)
        byt(9) = System.Text.Encoding.UTF8.GetBytes(pkm4_nickname.Text)
        byt(10) = System.Text.Encoding.UTF8.GetBytes(pkm5_nickname.Text)
        byt(11) = System.Text.Encoding.UTF8.GetBytes(pkm6_nickname.Text)
        Try
            Cursor = Cursors.WaitCursor
            Dim WriteString As String = _
                pkm1.SelectedIndex & "," & Convert.ToBase64String(byt(6)) & "," & _
                pkm1_helditem.SelectedIndex & "," & pkm1_gender.SelectedIndex & "," & _
                pkm1_ability.SelectedIndex & "," & pkm1_nature.SelectedIndex & "," & _
                pkm1_level.Value & "," & pkm1_iv1.Value & "," & _
                pkm1_iv2.Value & "," & pkm1_iv3.Value & "," & _
                pkm1_iv4.Value & "," & pkm1_iv5.Value & "," & _
                pkm1_iv6.Value & "," & pkm1_ev1.Value & "," & _
                pkm1_ev2.Value & "," & pkm1_ev3.Value & "," & _
                pkm1_ev4.Value & "," & pkm1_ev5.Value & "," & _
                pkm1_ev6.Value & "," & pkm1_shiny.CheckState.CompareTo(CheckState.Checked) & "," & _
                pkm1_move1.SelectedIndex & "," & pkm1_move2.SelectedIndex & "," & _
                pkm1_move3.SelectedIndex & "," & pkm1_move4.SelectedIndex & "," & _
                Convert.ToBase64String(byt(0)) & vbCrLf & _
                pkm2.SelectedIndex & "," & Convert.ToBase64String(byt(7)) & "," & _
                pkm2_helditem.SelectedIndex & "," & pkm2_gender.SelectedIndex & "," & _
                pkm2_ability.SelectedIndex & "," & pkm2_nature.SelectedIndex & "," & _
                pkm2_level.Value & "," & pkm2_iv1.Value & "," & _
                pkm2_iv2.Value & "," & pkm2_iv3.Value & "," & _
                pkm2_iv4.Value & "," & pkm2_iv5.Value & "," & _
                pkm2_iv6.Value & "," & pkm2_ev1.Value & "," & _
                pkm2_ev2.Value & "," & pkm2_ev3.Value & "," & _
                pkm2_ev4.Value & "," & pkm2_ev5.Value & "," & _
                pkm2_ev6.Value & "," & pkm2_shiny.CheckState.CompareTo(CheckState.Checked) & "," & _
                pkm2_move1.SelectedIndex & "," & pkm2_move2.SelectedIndex & "," & _
                pkm2_move3.SelectedIndex & "," & pkm2_move4.SelectedIndex & "," & _
                Convert.ToBase64String(byt(1)) & vbCrLf & _
                pkm3.SelectedIndex & "," & Convert.ToBase64String(byt(8)) & "," & _
                pkm3_helditem.SelectedIndex & "," & pkm3_gender.SelectedIndex & "," & _
                pkm3_ability.SelectedIndex & "," & pkm3_nature.SelectedIndex & "," & _
                pkm3_level.Value & "," & pkm3_iv1.Value & "," & _
                pkm3_iv2.Value & "," & pkm3_iv3.Value & "," & _
                pkm3_iv4.Value & "," & pkm3_iv5.Value & "," & _
                pkm3_iv6.Value & "," & pkm3_ev1.Value & "," & _
                pkm3_ev2.Value & "," & pkm3_ev3.Value & "," & _
                pkm3_ev4.Value & "," & pkm3_ev5.Value & "," & _
                pkm3_ev6.Value & "," & pkm3_shiny.CheckState.CompareTo(CheckState.Checked) & "," & _
                pkm3_move1.SelectedIndex & "," & pkm3_move2.SelectedIndex & "," & _
                pkm3_move3.SelectedIndex & "," & pkm3_move4.SelectedIndex & "," & _
                Convert.ToBase64String(byt(2)) & vbCrLf & _
                pkm4.SelectedIndex & "," & Convert.ToBase64String(byt(9)) & "," & _
                pkm4_helditem.SelectedIndex & "," & pkm4_gender.SelectedIndex & "," & _
                pkm4_ability.SelectedIndex & "," & pkm4_nature.SelectedIndex & "," & _
                pkm4_level.Value & "," & pkm4_iv1.Value & "," & _
                pkm4_iv2.Value & "," & pkm4_iv3.Value & "," & _
                pkm4_iv4.Value & "," & pkm4_iv5.Value & "," & _
                pkm4_iv6.Value & "," & pkm4_ev1.Value & "," & _
                pkm4_ev2.Value & "," & pkm4_ev3.Value & "," & _
                pkm4_ev4.Value & "," & pkm4_ev5.Value & "," & _
                pkm4_ev6.Value & "," & pkm4_shiny.CheckState.CompareTo(CheckState.Checked) & "," & _
                pkm4_move1.SelectedIndex & "," & pkm4_move2.SelectedIndex & "," & _
                pkm4_move3.SelectedIndex & "," & pkm4_move4.SelectedIndex & "," & _
                Convert.ToBase64String(byt(3)) & vbCrLf & _
                pkm5.SelectedIndex & "," & Convert.ToBase64String(byt(10)) & "," & _
                pkm5_helditem.SelectedIndex & "," & pkm5_gender.SelectedIndex & "," & _
                pkm5_ability.SelectedIndex & "," & pkm5_nature.SelectedIndex & "," & _
                pkm5_level.Value & "," & pkm5_iv1.Value & "," & _
                pkm5_iv2.Value & "," & pkm5_iv3.Value & "," & _
                pkm5_iv4.Value & "," & pkm5_iv5.Value & "," & _
                pkm5_iv6.Value & "," & pkm5_ev1.Value & "," & _
                pkm5_ev2.Value & "," & pkm5_ev3.Value & "," & _
                pkm5_ev4.Value & "," & pkm5_ev5.Value & "," & _
                pkm5_ev6.Value & "," & pkm5_shiny.CheckState.CompareTo(CheckState.Checked) & "," & _
                pkm5_move1.SelectedIndex & "," & pkm5_move2.SelectedIndex & "," & _
                pkm5_move3.SelectedIndex & "," & pkm5_move4.SelectedIndex & "," & _
                Convert.ToBase64String(byt(4)) & vbCrLf & _
                pkm6.SelectedIndex & "," & Convert.ToBase64String(byt(11)) & "," & _
                pkm6_helditem.SelectedIndex & "," & pkm6_gender.SelectedIndex & "," & _
                pkm6_ability.SelectedIndex & "," & pkm6_nature.SelectedIndex & "," & _
                pkm6_level.Value & "," & pkm6_iv1.Value & "," & _
                pkm6_iv2.Value & "," & pkm6_iv3.Value & "," & _
                pkm6_iv4.Value & "," & pkm6_iv5.Value & "," & _
                pkm6_iv6.Value & "," & pkm6_ev1.Value & "," & _
                pkm6_ev2.Value & "," & pkm6_ev3.Value & "," & _
                pkm6_ev4.Value & "," & pkm6_ev5.Value & "," & _
                pkm6_ev6.Value & "," & pkm6_shiny.CheckState.CompareTo(CheckState.Checked) & "," & _
                pkm6_move1.SelectedIndex & "," & pkm6_move2.SelectedIndex & "," & _
                pkm6_move3.SelectedIndex & "," & pkm6_move4.SelectedIndex & "," & _
                Convert.ToBase64String(byt(5))
            WriteString = WriteString & vbCrLf & (WriteString & "key").GetHashCode.ToString()

            Dim sw As New StreamWriter(filename, False, System.Text.Encoding.ASCII)
            sw.Write(WriteString)
            sw.Close()

            Me.Text = AppName & " - " & prettyFilename
            Cursor = Cursors.Default
        Catch ex As Exception
            Cursor = Cursors.Default
            MsgBox("The team could not be saved:" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Dim OpenName As String
        Dim PrettyOpen As String
        If OpenTeamDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            OpenName = OpenTeamDialog.FileName
            PrettyOpen = OpenName.Substring(OpenName.LastIndexOf("\") + 1, OpenName.Length - OpenName.LastIndexOf("\") - 1)
            If PrettyOpen.EndsWith(".tbt") Then PrettyOpen = PrettyOpen.Substring(0, PrettyOpen.Length - 4)
        Else
            Exit Sub
        End If
        Try
            Cursor = Cursors.WaitCursor
            Dim sr As New StreamReader(OpenName, System.Text.Encoding.ASCII)
            Dim set1 As String = sr.ReadLine()
            Dim set2 As String = sr.ReadLine()
            Dim set3 As String = sr.ReadLine()
            Dim set4 As String = sr.ReadLine()
            Dim set5 As String = sr.ReadLine()
            Dim set6 As String = sr.ReadLine()
            Dim hash As String = sr.ReadLine()
            sr.Close()
            Dim curindex As Integer

            Dim checksum As String = set1 & vbCrLf & _
                set2 & vbCrLf & set3 & vbCrLf & set4 & _
                vbCrLf & set5 & vbCrLf & set6
            If (checksum & "key").GetHashCode.ToString.Equals(hash) Then
                'Checksum passes. Open.
                Dim string1() As String = Split(set1, ",")
                Dim string2() As String = Split(set2, ",")
                Dim string3() As String = Split(set3, ",")
                Dim string4() As String = Split(set4, ",")
                Dim string5() As String = Split(set5, ",")
                Dim string6() As String = Split(set6, ",")

                curindex = Val(string1(0))
                pkm1.SelectedIndex = 0
                If curindex > 0 Then pkm1.SelectedIndex = curindex
                pkm1_nickname.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(1)))
                pkm1_helditem.SelectedIndex = Val(string1(2))
                pkm1_gender.SelectedIndex = Val(string1(3))
                pkm1_ability.SelectedIndex = Val(string1(4))
                pkm1_nature.SelectedIndex = Val(string1(5))
                pkm1_level.Value = Val(string1(6))
                pkm1_iv1.Value = Val(string1(7))
                pkm1_iv2.Value = Val(string1(8))
                pkm1_iv3.Value = Val(string1(9))
                pkm1_iv4.Value = Val(string1(10))
                pkm1_iv5.Value = Val(string1(11))
                pkm1_iv6.Value = Val(string1(12))
                pkm1_ev1.Value = Val(string1(13))
                pkm1_ev2.Value = Val(string1(14))
                pkm1_ev3.Value = Val(string1(15))
                pkm1_ev4.Value = Val(string1(16))
                pkm1_ev5.Value = Val(string1(17))
                pkm1_ev6.Value = Val(string1(18))
                pkm1_shiny.CheckState = CheckState.Unchecked
                If Val(string1(19)) = 0 Then pkm1_shiny.CheckState = CheckState.Checked
                pkm1_move1.SelectedIndex = Val(string1(20))
                pkm1_move2.SelectedIndex = Val(string1(21))
                pkm1_move3.SelectedIndex = Val(string1(22))
                pkm1_move4.SelectedIndex = Val(string1(23))
                pkm1_notes.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(24)))

                curindex = Val(string2(0))
                pkm2.SelectedIndex = 0
                If curindex > 0 Then pkm2.SelectedIndex = curindex
                pkm2_nickname.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string2(1)))
                pkm2_helditem.SelectedIndex = Val(string2(2))
                pkm2_gender.SelectedIndex = Val(string2(3))
                pkm2_ability.SelectedIndex = Val(string2(4))
                pkm2_nature.SelectedIndex = Val(string2(5))
                pkm2_level.Value = Val(string2(6))
                pkm2_iv1.Value = Val(string2(7))
                pkm2_iv2.Value = Val(string2(8))
                pkm2_iv3.Value = Val(string2(9))
                pkm2_iv4.Value = Val(string2(10))
                pkm2_iv5.Value = Val(string2(11))
                pkm2_iv6.Value = Val(string2(12))
                pkm2_ev1.Value = Val(string2(13))
                pkm2_ev2.Value = Val(string2(14))
                pkm2_ev3.Value = Val(string2(15))
                pkm2_ev4.Value = Val(string2(16))
                pkm2_ev5.Value = Val(string2(17))
                pkm2_ev6.Value = Val(string2(18))
                pkm2_shiny.CheckState = CheckState.Unchecked
                If Val(string2(19)) = 0 Then pkm2_shiny.CheckState = CheckState.Checked
                pkm2_move1.SelectedIndex = Val(string2(20))
                pkm2_move2.SelectedIndex = Val(string2(21))
                pkm2_move3.SelectedIndex = Val(string2(22))
                pkm2_move4.SelectedIndex = Val(string2(23))
                pkm2_notes.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string2(24)))

                curindex = Val(string3(0))
                pkm3.SelectedIndex = 0
                If curindex > 0 Then pkm3.SelectedIndex = curindex
                pkm3_nickname.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string3(1)))
                pkm3_helditem.SelectedIndex = Val(string3(2))
                pkm3_gender.SelectedIndex = Val(string3(3))
                pkm3_ability.SelectedIndex = Val(string3(4))
                pkm3_nature.SelectedIndex = Val(string3(5))
                pkm3_level.Value = Val(string3(6))
                pkm3_iv1.Value = Val(string3(7))
                pkm3_iv2.Value = Val(string3(8))
                pkm3_iv3.Value = Val(string3(9))
                pkm3_iv4.Value = Val(string3(10))
                pkm3_iv5.Value = Val(string3(11))
                pkm3_iv6.Value = Val(string3(12))
                pkm3_ev1.Value = Val(string3(13))
                pkm3_ev2.Value = Val(string3(14))
                pkm3_ev3.Value = Val(string3(15))
                pkm3_ev4.Value = Val(string3(16))
                pkm3_ev5.Value = Val(string3(17))
                pkm3_ev6.Value = Val(string3(18))
                pkm3_shiny.CheckState = CheckState.Unchecked
                If Val(string3(19)) = 0 Then pkm3_shiny.CheckState = CheckState.Checked
                pkm3_move1.SelectedIndex = Val(string3(20))
                pkm3_move2.SelectedIndex = Val(string3(21))
                pkm3_move3.SelectedIndex = Val(string3(22))
                pkm3_move4.SelectedIndex = Val(string3(23))
                pkm3_notes.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string3(24)))

                curindex = Val(string4(0))
                pkm4.SelectedIndex = 0
                If curindex > 0 Then pkm4.SelectedIndex = curindex
                pkm4_nickname.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string4(1)))
                pkm4_helditem.SelectedIndex = Val(string4(2))
                pkm4_gender.SelectedIndex = Val(string4(3))
                pkm4_ability.SelectedIndex = Val(string4(4))
                pkm4_nature.SelectedIndex = Val(string4(5))
                pkm4_level.Value = Val(string4(6))
                pkm4_iv1.Value = Val(string4(7))
                pkm4_iv2.Value = Val(string4(8))
                pkm4_iv3.Value = Val(string4(9))
                pkm4_iv4.Value = Val(string4(10))
                pkm4_iv5.Value = Val(string4(11))
                pkm4_iv6.Value = Val(string4(12))
                pkm4_ev1.Value = Val(string4(13))
                pkm4_ev2.Value = Val(string4(14))
                pkm4_ev3.Value = Val(string4(15))
                pkm4_ev4.Value = Val(string4(16))
                pkm4_ev5.Value = Val(string4(17))
                pkm4_ev6.Value = Val(string4(18))
                pkm4_shiny.CheckState = CheckState.Unchecked
                If Val(string4(19)) = 0 Then pkm4_shiny.CheckState = CheckState.Checked
                pkm4_move1.SelectedIndex = Val(string4(20))
                pkm4_move2.SelectedIndex = Val(string4(21))
                pkm4_move3.SelectedIndex = Val(string4(22))
                pkm4_move4.SelectedIndex = Val(string4(23))
                pkm4_notes.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string4(24)))

                curindex = Val(string5(0))
                pkm5.SelectedIndex = 0
                If curindex > 0 Then pkm5.SelectedIndex = curindex
                pkm5_nickname.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string5(1)))
                pkm5_helditem.SelectedIndex = Val(string5(2))
                pkm5_gender.SelectedIndex = Val(string5(3))
                pkm5_ability.SelectedIndex = Val(string5(4))
                pkm5_nature.SelectedIndex = Val(string5(5))
                pkm5_level.Value = Val(string5(6))
                pkm5_iv1.Value = Val(string5(7))
                pkm5_iv2.Value = Val(string5(8))
                pkm5_iv3.Value = Val(string5(9))
                pkm5_iv4.Value = Val(string5(10))
                pkm5_iv5.Value = Val(string5(11))
                pkm5_iv6.Value = Val(string5(12))
                pkm5_ev1.Value = Val(string5(13))
                pkm5_ev2.Value = Val(string5(14))
                pkm5_ev3.Value = Val(string5(15))
                pkm5_ev4.Value = Val(string5(16))
                pkm5_ev5.Value = Val(string5(17))
                pkm5_ev6.Value = Val(string5(18))
                pkm5_shiny.CheckState = CheckState.Unchecked
                If Val(string5(19)) = 0 Then pkm5_shiny.CheckState = CheckState.Checked
                pkm5_move1.SelectedIndex = Val(string5(20))
                pkm5_move2.SelectedIndex = Val(string5(21))
                pkm5_move3.SelectedIndex = Val(string5(22))
                pkm5_move4.SelectedIndex = Val(string5(23))
                pkm5_notes.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string5(24)))

                curindex = Val(string6(0))
                pkm6.SelectedIndex = 0
                If curindex > 0 Then pkm6.SelectedIndex = curindex
                pkm6_nickname.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string6(1)))
                pkm6_helditem.SelectedIndex = Val(string6(2))
                pkm6_gender.SelectedIndex = Val(string6(3))
                pkm6_ability.SelectedIndex = Val(string6(4))
                pkm6_nature.SelectedIndex = Val(string6(5))
                pkm6_level.Value = Val(string6(6))
                pkm6_iv1.Value = Val(string6(7))
                pkm6_iv2.Value = Val(string6(8))
                pkm6_iv3.Value = Val(string6(9))
                pkm6_iv4.Value = Val(string6(10))
                pkm6_iv5.Value = Val(string6(11))
                pkm6_iv6.Value = Val(string6(12))
                pkm6_ev1.Value = Val(string6(13))
                pkm6_ev2.Value = Val(string6(14))
                pkm6_ev3.Value = Val(string6(15))
                pkm6_ev4.Value = Val(string6(16))
                pkm6_ev5.Value = Val(string6(17))
                pkm6_ev6.Value = Val(string6(18))
                pkm6_shiny.CheckState = CheckState.Unchecked
                If Val(string6(19)) = 0 Then pkm6_shiny.CheckState = CheckState.Checked
                pkm6_move1.SelectedIndex = Val(string6(20))
                pkm6_move2.SelectedIndex = Val(string6(21))
                pkm6_move3.SelectedIndex = Val(string6(22))
                pkm6_move4.SelectedIndex = Val(string6(23))
                pkm6_notes.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string6(24)))


            Else
                Cursor = Cursors.Default
                MsgBox("The team is corrupt.", MsgBoxStyle.Critical)
                Exit Sub
            End If
            filename = OpenName
            prettyFilename = PrettyOpen
            Me.Text = AppName & " - " & prettyFilename
            Cursor = Cursors.Default
            If Tabs.SelectedIndex = 6 Then Analyze()
        Catch ex As Exception
            Cursor = Cursors.Default
            MsgBox("The team could not be opened:" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub HelpToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem1.Click
        If Not Help.Visible Then Help.Show()
    End Sub

    Public Function GetOneFilename()
        If SaveMonDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            Return SaveMonDialog.FileName
        Else
            Return ""
        End If
    End Function

    Private Sub SavePokToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SavePokToolStripMenuItem.Click
        Dim OneName As String = ""
        Dim SaveString As String = ""
        Cursor = Cursors.WaitCursor
        If Tabs.SelectedIndex = 0 And pkm1.SelectedIndex > 0 Then
            OneName = GetOneFilename()
            If OneName.Length = 0 Then
                Cursor = Cursors.Default
                Exit Sub
            End If
            SaveString = pkm1.SelectedIndex & "," & Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(pkm1_nickname.Text)) & "," & _
                            pkm1_helditem.SelectedIndex & "," & pkm1_gender.SelectedIndex & "," & _
                            pkm1_ability.SelectedIndex & "," & pkm1_nature.SelectedIndex & "," & _
                            pkm1_level.Value & "," & pkm1_iv1.Value & "," & _
                            pkm1_iv2.Value & "," & pkm1_iv3.Value & "," & _
                            pkm1_iv4.Value & "," & pkm1_iv5.Value & "," & _
                            pkm1_iv6.Value & "," & pkm1_ev1.Value & "," & _
                            pkm1_ev2.Value & "," & pkm1_ev3.Value & "," & _
                            pkm1_ev4.Value & "," & pkm1_ev5.Value & "," & _
                            pkm1_ev6.Value & "," & pkm1_shiny.CheckState.CompareTo(CheckState.Checked) & "," & _
                            pkm1_move1.SelectedIndex & "," & pkm1_move2.SelectedIndex & "," & _
                            pkm1_move3.SelectedIndex & "," & pkm1_move4.SelectedIndex & "," & _
                            Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(pkm1_notes.Text))
        ElseIf Tabs.SelectedIndex = 1 And pkm2.SelectedIndex > 0 Then
            OneName = GetOneFilename()
            If OneName.Length = 0 Then
                Cursor = Cursors.Default
                Exit Sub
            End If
            SaveString = pkm2.SelectedIndex & "," & Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(pkm2_nickname.Text)) & "," & _
                            pkm2_helditem.SelectedIndex & "," & pkm2_gender.SelectedIndex & "," & _
                            pkm2_ability.SelectedIndex & "," & pkm2_nature.SelectedIndex & "," & _
                            pkm2_level.Value & "," & pkm2_iv1.Value & "," & _
                            pkm2_iv2.Value & "," & pkm2_iv3.Value & "," & _
                            pkm2_iv4.Value & "," & pkm2_iv5.Value & "," & _
                            pkm2_iv6.Value & "," & pkm2_ev1.Value & "," & _
                            pkm2_ev2.Value & "," & pkm2_ev3.Value & "," & _
                            pkm2_ev4.Value & "," & pkm2_ev5.Value & "," & _
                            pkm2_ev6.Value & "," & pkm2_shiny.CheckState.CompareTo(CheckState.Checked) & "," & _
                            pkm2_move1.SelectedIndex & "," & pkm2_move2.SelectedIndex & "," & _
                            pkm2_move3.SelectedIndex & "," & pkm2_move4.SelectedIndex & "," & _
                            Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(pkm2_notes.Text))
        ElseIf Tabs.SelectedIndex = 2 And pkm3.SelectedIndex > 0 Then
            OneName = GetOneFilename()
            If OneName.Length = 0 Then
                Cursor = Cursors.Default
                Exit Sub
            End If
            SaveString = pkm3.SelectedIndex & "," & Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(pkm3_nickname.Text)) & "," & _
                                pkm3_helditem.SelectedIndex & "," & pkm3_gender.SelectedIndex & "," & _
                                pkm3_ability.SelectedIndex & "," & pkm3_nature.SelectedIndex & "," & _
                                pkm3_level.Value & "," & pkm3_iv1.Value & "," & _
                                pkm3_iv2.Value & "," & pkm3_iv3.Value & "," & _
                                pkm3_iv4.Value & "," & pkm3_iv5.Value & "," & _
                                pkm3_iv6.Value & "," & pkm3_ev1.Value & "," & _
                                pkm3_ev2.Value & "," & pkm3_ev3.Value & "," & _
                                pkm3_ev4.Value & "," & pkm3_ev5.Value & "," & _
                                pkm3_ev6.Value & "," & pkm3_shiny.CheckState.CompareTo(CheckState.Checked) & "," & _
                                pkm3_move1.SelectedIndex & "," & pkm3_move2.SelectedIndex & "," & _
                                pkm3_move3.SelectedIndex & "," & pkm3_move4.SelectedIndex & "," & _
                                Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(pkm3_notes.Text))
        ElseIf Tabs.SelectedIndex = 3 And pkm4.SelectedIndex > 0 Then
            OneName = GetOneFilename()
            If OneName.Length = 0 Then
                Cursor = Cursors.Default
                Exit Sub
            End If
            SaveString = pkm4.SelectedIndex & "," & Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(pkm4_nickname.Text)) & "," & _
                            pkm4_helditem.SelectedIndex & "," & pkm4_gender.SelectedIndex & "," & _
                            pkm4_ability.SelectedIndex & "," & pkm4_nature.SelectedIndex & "," & _
                            pkm4_level.Value & "," & pkm4_iv1.Value & "," & _
                            pkm4_iv2.Value & "," & pkm4_iv3.Value & "," & _
                            pkm4_iv4.Value & "," & pkm4_iv5.Value & "," & _
                            pkm4_iv6.Value & "," & pkm4_ev1.Value & "," & _
                            pkm4_ev2.Value & "," & pkm4_ev3.Value & "," & _
                            pkm4_ev4.Value & "," & pkm4_ev5.Value & "," & _
                            pkm4_ev6.Value & "," & pkm4_shiny.CheckState.CompareTo(CheckState.Checked) & "," & _
                            pkm4_move1.SelectedIndex & "," & pkm4_move2.SelectedIndex & "," & _
                            pkm4_move3.SelectedIndex & "," & pkm4_move4.SelectedIndex & "," & _
                            Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(pkm4_notes.Text))
        ElseIf Tabs.SelectedIndex = 4 And pkm5.SelectedIndex > 0 Then
            OneName = GetOneFilename()
            If OneName.Length = 0 Then
                Cursor = Cursors.Default
                Exit Sub
            End If
            SaveString = pkm5.SelectedIndex & "," & Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(pkm5_nickname.Text)) & "," & _
                            pkm5_helditem.SelectedIndex & "," & pkm5_gender.SelectedIndex & "," & _
                            pkm5_ability.SelectedIndex & "," & pkm5_nature.SelectedIndex & "," & _
                            pkm5_level.Value & "," & pkm5_iv1.Value & "," & _
                            pkm5_iv2.Value & "," & pkm5_iv3.Value & "," & _
                            pkm5_iv4.Value & "," & pkm5_iv5.Value & "," & _
                            pkm5_iv6.Value & "," & pkm5_ev1.Value & "," & _
                            pkm5_ev2.Value & "," & pkm5_ev3.Value & "," & _
                            pkm5_ev4.Value & "," & pkm5_ev5.Value & "," & _
                            pkm5_ev6.Value & "," & pkm5_shiny.CheckState.CompareTo(CheckState.Checked) & "," & _
                            pkm5_move1.SelectedIndex & "," & pkm5_move2.SelectedIndex & "," & _
                            pkm5_move3.SelectedIndex & "," & pkm5_move4.SelectedIndex & "," & _
                            Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(pkm5_notes.Text))
        ElseIf Tabs.SelectedIndex = 5 And pkm6.SelectedIndex > 0 Then
            OneName = GetOneFilename()
            If OneName.Length = 0 Then
                Cursor = Cursors.Default
                Exit Sub
            End If
            SaveString = pkm6.SelectedIndex & "," & Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(pkm6_nickname.Text)) & "," & _
                            pkm6_helditem.SelectedIndex & "," & pkm6_gender.SelectedIndex & "," & _
                            pkm6_ability.SelectedIndex & "," & pkm6_nature.SelectedIndex & "," & _
                            pkm6_level.Value & "," & pkm6_iv1.Value & "," & _
                            pkm6_iv2.Value & "," & pkm6_iv3.Value & "," & _
                            pkm6_iv4.Value & "," & pkm6_iv5.Value & "," & _
                            pkm6_iv6.Value & "," & pkm6_ev1.Value & "," & _
                            pkm6_ev2.Value & "," & pkm6_ev3.Value & "," & _
                            pkm6_ev4.Value & "," & pkm6_ev5.Value & "," & _
                            pkm6_ev6.Value & "," & pkm6_shiny.CheckState.CompareTo(CheckState.Checked) & "," & _
                            pkm6_move1.SelectedIndex & "," & pkm6_move2.SelectedIndex & "," & _
                            pkm6_move3.SelectedIndex & "," & pkm6_move4.SelectedIndex & "," & _
                            Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(pkm6_notes.Text))
        Else
            Cursor = Cursors.Default
            MsgBox("You do not have a Pokémon selected.", MsgBoxStyle.Exclamation)
        End If
        If SaveString.Length > 0 Then
            SaveString = SaveString & vbCrLf & (SaveString & "key").GetHashCode.ToString()
            Try
                Dim sw As New StreamWriter(OneName, False, System.Text.Encoding.ASCII)
                sw.Write(SaveString)
                sw.Close()
            Catch ex As Exception
                Cursor = Cursors.Default
                MsgBox("The Pokémon could not be saved:" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
        Cursor = Cursors.Default
    End Sub

    Private Sub LoadPokémonToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadPokémonToolStripMenuItem.Click
        If Tabs.SelectedIndex < 6 Then
            Dim OpenName As String
            If OpenMonDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                OpenName = OpenMonDialog.FileName
            Else
                Exit Sub
            End If
            Try
                Cursor = Cursors.WaitCursor
                Dim sr As New StreamReader(OpenName, System.Text.Encoding.ASCII)
                Dim set1 As String = sr.ReadLine()
                Dim hash As String = sr.ReadLine()
                sr.Close()

                Dim checksum As String = set1
                If (checksum & "key").GetHashCode.ToString.Equals(hash) Then
                    'Checksum passes. Open.
                    Dim string1() As String = Split(set1, ",")
                    If Tabs.SelectedIndex = 0 Then
                        pkm1.SelectedIndex = Val(string1(0))
                        pkm1_nickname.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(1)))
                        pkm1_helditem.SelectedIndex = Val(string1(2))
                        pkm1_gender.SelectedIndex = Val(string1(3))
                        pkm1_ability.SelectedIndex = Val(string1(4))
                        pkm1_nature.SelectedIndex = Val(string1(5))
                        pkm1_level.Value = Val(string1(6))
                        pkm1_iv1.Value = Val(string1(7))
                        pkm1_iv2.Value = Val(string1(8))
                        pkm1_iv3.Value = Val(string1(9))
                        pkm1_iv4.Value = Val(string1(10))
                        pkm1_iv5.Value = Val(string1(11))
                        pkm1_iv6.Value = Val(string1(12))
                        pkm1_ev1.Value = Val(string1(13))
                        pkm1_ev2.Value = Val(string1(14))
                        pkm1_ev3.Value = Val(string1(15))
                        pkm1_ev4.Value = Val(string1(16))
                        pkm1_ev5.Value = Val(string1(17))
                        pkm1_ev6.Value = Val(string1(18))
                        pkm1_shiny.CheckState = CheckState.Unchecked
                        If Val(string1(19)) = 0 Then pkm1_shiny.CheckState = CheckState.Checked
                        pkm1_move1.SelectedIndex = Val(string1(20))
                        pkm1_move2.SelectedIndex = Val(string1(21))
                        pkm1_move3.SelectedIndex = Val(string1(22))
                        pkm1_move4.SelectedIndex = Val(string1(23))
                        pkm1_notes.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(24)))
                    ElseIf Tabs.SelectedIndex = 1 Then
                        pkm2.SelectedIndex = Val(string1(0))
                        pkm2_nickname.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(1)))
                        pkm2_helditem.SelectedIndex = Val(string1(2))
                        pkm2_gender.SelectedIndex = Val(string1(3))
                        pkm2_ability.SelectedIndex = Val(string1(4))
                        pkm2_nature.SelectedIndex = Val(string1(5))
                        pkm2_level.Value = Val(string1(6))
                        pkm2_iv1.Value = Val(string1(7))
                        pkm2_iv2.Value = Val(string1(8))
                        pkm2_iv3.Value = Val(string1(9))
                        pkm2_iv4.Value = Val(string1(10))
                        pkm2_iv5.Value = Val(string1(11))
                        pkm2_iv6.Value = Val(string1(12))
                        pkm2_ev1.Value = Val(string1(13))
                        pkm2_ev2.Value = Val(string1(14))
                        pkm2_ev3.Value = Val(string1(15))
                        pkm2_ev4.Value = Val(string1(16))
                        pkm2_ev5.Value = Val(string1(17))
                        pkm2_ev6.Value = Val(string1(18))
                        pkm2_shiny.CheckState = CheckState.Unchecked
                        If Val(string1(19)) = 0 Then pkm2_shiny.CheckState = CheckState.Checked
                        pkm2_move1.SelectedIndex = Val(string1(20))
                        pkm2_move2.SelectedIndex = Val(string1(21))
                        pkm2_move3.SelectedIndex = Val(string1(22))
                        pkm2_move4.SelectedIndex = Val(string1(23))
                        pkm2_notes.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(24)))
                    ElseIf Tabs.SelectedIndex = 2 Then
                        pkm3.SelectedIndex = Val(string1(0))
                        pkm3_nickname.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(1)))
                        pkm3_helditem.SelectedIndex = Val(string1(2))
                        pkm3_gender.SelectedIndex = Val(string1(3))
                        pkm3_ability.SelectedIndex = Val(string1(4))
                        pkm3_nature.SelectedIndex = Val(string1(5))
                        pkm3_level.Value = Val(string1(6))
                        pkm3_iv1.Value = Val(string1(7))
                        pkm3_iv2.Value = Val(string1(8))
                        pkm3_iv3.Value = Val(string1(9))
                        pkm3_iv4.Value = Val(string1(10))
                        pkm3_iv5.Value = Val(string1(11))
                        pkm3_iv6.Value = Val(string1(12))
                        pkm3_ev1.Value = Val(string1(13))
                        pkm3_ev2.Value = Val(string1(14))
                        pkm3_ev3.Value = Val(string1(15))
                        pkm3_ev4.Value = Val(string1(16))
                        pkm3_ev5.Value = Val(string1(17))
                        pkm3_ev6.Value = Val(string1(18))
                        pkm3_shiny.CheckState = CheckState.Unchecked
                        If Val(string1(19)) = 0 Then pkm3_shiny.CheckState = CheckState.Checked
                        pkm3_move1.SelectedIndex = Val(string1(20))
                        pkm3_move2.SelectedIndex = Val(string1(21))
                        pkm3_move3.SelectedIndex = Val(string1(22))
                        pkm3_move4.SelectedIndex = Val(string1(23))
                        pkm3_notes.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(24)))
                    ElseIf Tabs.SelectedIndex = 3 Then
                        pkm4.SelectedIndex = Val(string1(0))
                        pkm4_nickname.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(1)))
                        pkm4_helditem.SelectedIndex = Val(string1(2))
                        pkm4_gender.SelectedIndex = Val(string1(3))
                        pkm4_ability.SelectedIndex = Val(string1(4))
                        pkm4_nature.SelectedIndex = Val(string1(5))
                        pkm4_level.Value = Val(string1(6))
                        pkm4_iv1.Value = Val(string1(7))
                        pkm4_iv2.Value = Val(string1(8))
                        pkm4_iv3.Value = Val(string1(9))
                        pkm4_iv4.Value = Val(string1(10))
                        pkm4_iv5.Value = Val(string1(11))
                        pkm4_iv6.Value = Val(string1(12))
                        pkm4_ev1.Value = Val(string1(13))
                        pkm4_ev2.Value = Val(string1(14))
                        pkm4_ev3.Value = Val(string1(15))
                        pkm4_ev4.Value = Val(string1(16))
                        pkm4_ev5.Value = Val(string1(17))
                        pkm4_ev6.Value = Val(string1(18))
                        pkm4_shiny.CheckState = CheckState.Unchecked
                        If Val(string1(19)) = 0 Then pkm4_shiny.CheckState = CheckState.Checked
                        pkm4_move1.SelectedIndex = Val(string1(20))
                        pkm4_move2.SelectedIndex = Val(string1(21))
                        pkm4_move3.SelectedIndex = Val(string1(22))
                        pkm4_move4.SelectedIndex = Val(string1(23))
                        pkm4_notes.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(24)))
                    ElseIf Tabs.SelectedIndex = 4 Then
                        pkm5.SelectedIndex = Val(string1(0))
                        pkm5_nickname.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(1)))
                        pkm5_helditem.SelectedIndex = Val(string1(2))
                        pkm5_gender.SelectedIndex = Val(string1(3))
                        pkm5_ability.SelectedIndex = Val(string1(4))
                        pkm5_nature.SelectedIndex = Val(string1(5))
                        pkm5_level.Value = Val(string1(6))
                        pkm5_iv1.Value = Val(string1(7))
                        pkm5_iv2.Value = Val(string1(8))
                        pkm5_iv3.Value = Val(string1(9))
                        pkm5_iv4.Value = Val(string1(10))
                        pkm5_iv5.Value = Val(string1(11))
                        pkm5_iv6.Value = Val(string1(12))
                        pkm5_ev1.Value = Val(string1(13))
                        pkm5_ev2.Value = Val(string1(14))
                        pkm5_ev3.Value = Val(string1(15))
                        pkm5_ev4.Value = Val(string1(16))
                        pkm5_ev5.Value = Val(string1(17))
                        pkm5_ev6.Value = Val(string1(18))
                        pkm5_shiny.CheckState = CheckState.Unchecked
                        If Val(string1(19)) = 0 Then pkm5_shiny.CheckState = CheckState.Checked
                        pkm5_move1.SelectedIndex = Val(string1(20))
                        pkm5_move2.SelectedIndex = Val(string1(21))
                        pkm5_move3.SelectedIndex = Val(string1(22))
                        pkm5_move4.SelectedIndex = Val(string1(23))
                        pkm5_notes.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(24)))
                    ElseIf Tabs.SelectedIndex = 5 Then
                        pkm6.SelectedIndex = Val(string1(0))
                        pkm6_nickname.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(1)))
                        pkm6_helditem.SelectedIndex = Val(string1(2))
                        pkm6_gender.SelectedIndex = Val(string1(3))
                        pkm6_ability.SelectedIndex = Val(string1(4))
                        pkm6_nature.SelectedIndex = Val(string1(5))
                        pkm6_level.Value = Val(string1(6))
                        pkm6_iv1.Value = Val(string1(7))
                        pkm6_iv2.Value = Val(string1(8))
                        pkm6_iv3.Value = Val(string1(9))
                        pkm6_iv4.Value = Val(string1(10))
                        pkm6_iv5.Value = Val(string1(11))
                        pkm6_iv6.Value = Val(string1(12))
                        pkm6_ev1.Value = Val(string1(13))
                        pkm6_ev2.Value = Val(string1(14))
                        pkm6_ev3.Value = Val(string1(15))
                        pkm6_ev4.Value = Val(string1(16))
                        pkm6_ev5.Value = Val(string1(17))
                        pkm6_ev6.Value = Val(string1(18))
                        pkm6_shiny.CheckState = CheckState.Unchecked
                        If Val(string1(19)) = 0 Then pkm6_shiny.CheckState = CheckState.Checked
                        pkm6_move1.SelectedIndex = Val(string1(20))
                        pkm6_move2.SelectedIndex = Val(string1(21))
                        pkm6_move3.SelectedIndex = Val(string1(22))
                        pkm6_move4.SelectedIndex = Val(string1(23))
                        pkm6_notes.Text = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(string1(24)))
                    Else
                        Cursor = Cursors.Default
                        MsgBox("This should not happen.", MsgBoxStyle.Critical)
                        Exit Sub
                    End If
                Else
                    Cursor = Cursors.Default
                    MsgBox("The Pokémon is corrupt.", MsgBoxStyle.Critical)
                    Exit Sub
                End If
                Cursor = Cursors.Default
            Catch ex As Exception
                Cursor = Cursors.Default
                MsgBox("The Pokémon could not be imported:" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical)
            End Try
        Else
            Cursor = Cursors.Default
            MsgBox("You do not have a Pokémon selected.", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Sub DeletePokémonToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeletePokémonToolStripMenuItem.Click
        If Tabs.SelectedIndex = 0 Then
            If MsgBox("Are you sure you want to erase this Pokémon?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                pkm1.SelectedIndex = 0
                pkm1_nickname.Text = ""
                pkm1_helditem.SelectedIndex = 0
                pkm1_gender.SelectedIndex = 0
                pkm1_nature.SelectedIndex = 0
                pkm1_level.Value = 100
                pkm1_iv1.Value = 0
                pkm1_iv2.Value = 0
                pkm1_iv3.Value = 0
                pkm1_iv4.Value = 0
                pkm1_iv5.Value = 0
                pkm1_iv6.Value = 0
                pkm1_ev1.Value = 0
                pkm1_ev2.Value = 0
                pkm1_ev3.Value = 0
                pkm1_ev4.Value = 0
                pkm1_ev5.Value = 0
                pkm1_ev6.Value = 0
                pkm1_shiny.CheckState = CheckState.Unchecked
                pkm1_move1.SelectedIndex = 0
                pkm1_move2.SelectedIndex = 0
                pkm1_move3.SelectedIndex = 0
                pkm1_move4.SelectedIndex = 0
                pkm1_notes.Text = ""
            End If
        ElseIf Tabs.SelectedIndex = 1 Then
            If MsgBox("Are you sure you want to erase this Pokémon?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                pkm2.SelectedIndex = 0
                pkm2_nickname.Text = ""
                pkm2_helditem.SelectedIndex = 0
                pkm2_gender.SelectedIndex = 0
                pkm2_nature.SelectedIndex = 0
                pkm2_level.Value = 100
                pkm2_iv1.Value = 0
                pkm2_iv2.Value = 0
                pkm2_iv3.Value = 0
                pkm2_iv4.Value = 0
                pkm2_iv5.Value = 0
                pkm2_iv6.Value = 0
                pkm2_ev1.Value = 0
                pkm2_ev2.Value = 0
                pkm2_ev3.Value = 0
                pkm2_ev4.Value = 0
                pkm2_ev5.Value = 0
                pkm2_ev6.Value = 0
                pkm2_shiny.CheckState = CheckState.Unchecked
                pkm2_move1.SelectedIndex = 0
                pkm2_move2.SelectedIndex = 0
                pkm2_move3.SelectedIndex = 0
                pkm2_move4.SelectedIndex = 0
                pkm2_notes.Text = ""
            End If
        ElseIf Tabs.SelectedIndex = 2 Then
            If MsgBox("Are you sure you want to erase this Pokémon?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                pkm3.SelectedIndex = 0
                pkm3_nickname.Text = ""
                pkm3_helditem.SelectedIndex = 0
                pkm3_gender.SelectedIndex = 0
                pkm3_nature.SelectedIndex = 0
                pkm3_level.Value = 100
                pkm3_iv1.Value = 0
                pkm3_iv2.Value = 0
                pkm3_iv3.Value = 0
                pkm3_iv4.Value = 0
                pkm3_iv5.Value = 0
                pkm3_iv6.Value = 0
                pkm3_ev1.Value = 0
                pkm3_ev2.Value = 0
                pkm3_ev3.Value = 0
                pkm3_ev4.Value = 0
                pkm3_ev5.Value = 0
                pkm3_ev6.Value = 0
                pkm3_shiny.CheckState = CheckState.Unchecked
                pkm3_move1.SelectedIndex = 0
                pkm3_move2.SelectedIndex = 0
                pkm3_move3.SelectedIndex = 0
                pkm3_move4.SelectedIndex = 0
                pkm3_notes.Text = ""
            End If
        ElseIf Tabs.SelectedIndex = 3 Then
            If MsgBox("Are you sure you want to erase this Pokémon?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                pkm4.SelectedIndex = 0
                pkm4_nickname.Text = ""
                pkm4_helditem.SelectedIndex = 0
                pkm4_gender.SelectedIndex = 0
                pkm4_nature.SelectedIndex = 0
                pkm4_level.Value = 100
                pkm4_iv1.Value = 0
                pkm4_iv2.Value = 0
                pkm4_iv3.Value = 0
                pkm4_iv4.Value = 0
                pkm4_iv5.Value = 0
                pkm4_iv6.Value = 0
                pkm4_ev1.Value = 0
                pkm4_ev2.Value = 0
                pkm4_ev3.Value = 0
                pkm4_ev4.Value = 0
                pkm4_ev5.Value = 0
                pkm4_ev6.Value = 0
                pkm4_shiny.CheckState = CheckState.Unchecked
                pkm4_move1.SelectedIndex = 0
                pkm4_move2.SelectedIndex = 0
                pkm4_move3.SelectedIndex = 0
                pkm4_move4.SelectedIndex = 0
                pkm4_notes.Text = ""
            End If
        ElseIf Tabs.SelectedIndex = 4 Then
            If MsgBox("Are you sure you want to erase this Pokémon?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                pkm5.SelectedIndex = 0
                pkm5_nickname.Text = ""
                pkm5_helditem.SelectedIndex = 0
                pkm5_gender.SelectedIndex = 0
                pkm5_nature.SelectedIndex = 0
                pkm5_level.Value = 100
                pkm5_iv1.Value = 0
                pkm5_iv2.Value = 0
                pkm5_iv3.Value = 0
                pkm5_iv4.Value = 0
                pkm5_iv5.Value = 0
                pkm5_iv6.Value = 0
                pkm5_ev1.Value = 0
                pkm5_ev2.Value = 0
                pkm5_ev3.Value = 0
                pkm5_ev4.Value = 0
                pkm5_ev5.Value = 0
                pkm5_ev6.Value = 0
                pkm5_shiny.CheckState = CheckState.Unchecked
                pkm5_move1.SelectedIndex = 0
                pkm5_move2.SelectedIndex = 0
                pkm5_move3.SelectedIndex = 0
                pkm5_move4.SelectedIndex = 0
                pkm5_notes.Text = ""
            End If
        ElseIf Tabs.SelectedIndex = 5 Then
            If MsgBox("Are you sure you want to erase this Pokémon?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                pkm6.SelectedIndex = 0
                pkm6_nickname.Text = ""
                pkm6_helditem.SelectedIndex = 0
                pkm6_gender.SelectedIndex = 0
                pkm6_nature.SelectedIndex = 0
                pkm6_level.Value = 100
                pkm6_iv1.Value = 0
                pkm6_iv2.Value = 0
                pkm6_iv3.Value = 0
                pkm6_iv4.Value = 0
                pkm6_iv5.Value = 0
                pkm6_iv6.Value = 0
                pkm6_ev1.Value = 0
                pkm6_ev2.Value = 0
                pkm6_ev3.Value = 0
                pkm6_ev4.Value = 0
                pkm6_ev5.Value = 0
                pkm6_ev6.Value = 0
                pkm6_shiny.CheckState = CheckState.Unchecked
                pkm6_move1.SelectedIndex = 0
                pkm6_move2.SelectedIndex = 0
                pkm6_move3.SelectedIndex = 0
                pkm6_move4.SelectedIndex = 0
                pkm6_notes.Text = ""
            End If
        ElseIf Tabs.SelectedIndex = 6 Then
            MsgBox("You do not have a Pokémon selected.", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Sub TabPage7_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage7.Leave
        LoadPokémonToolStripMenuItem.Enabled = True
        SavePokToolStripMenuItem.Enabled = True
        DeletePokémonToolStripMenuItem.Enabled = True
    End Sub

    Private Sub SaveAsCSVToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsCSVToolStripMenuItem.Click
        If SaveCSVDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            DoSaveCSV(SaveCSVDialog.FileName)
        End If
    End Sub


    Public Sub DoSaveCSV(ByRef CSVFilename As String)
        Dim shinies() As String = {"", "", "", "", "", ""}
        If pkm1_shiny.CheckState = CheckState.Checked Then
            shinies(0) = "Shiny"
        End If
        If pkm2_shiny.CheckState = CheckState.Checked Then
            shinies(1) = "Shiny"
        End If
        If pkm3_shiny.CheckState = CheckState.Checked Then
            shinies(2) = "Shiny"
        End If
        If pkm4_shiny.CheckState = CheckState.Checked Then
            shinies(3) = "Shiny"
        End If
        If pkm5_shiny.CheckState = CheckState.Checked Then
            shinies(4) = "Shiny"
        End If
        If pkm6_shiny.CheckState = CheckState.Checked Then
            shinies(5) = "Shiny"
        End If
        
        Try
            Cursor = Cursors.WaitCursor
            Dim WriteString As String = _
                """Pokémon Species"",""Nickname"",""Held Item"",""Gender""," & _
                """Ability"",""Nature"",""Level""," & _
                """HP IV"",""Atk IV"",""Def IV"",""SpA IV"",""SpD IV"",""Spd IV""," & _
                """HP EV"",""Atk EV"",""Def EV"",""SpA EV"",""SpD EV"",""Spd EV""," & _
                """HP"",""Attack"",""Defense"",""Sp. Atk"",""Sp. Def"",""Speed""," & _
                """Shiny"",""Move 1"",""Move 2"",""Move 3"",""Move 4"",""Notes""" & vbCrLf & _
                """" & pkm1.Text & """,""" & pkm1_nickname.Text & """,""" & _
                pkm1_helditem.Text & """,""" & pkm1_gender.Text & """,""" & _
                pkm1_ability.Text & """,""" & pkm1_nature.Text & """,""" & _
                pkm1_level.Value & """,""" & pkm1_iv1.Value & """,""" & _
                pkm1_iv2.Value & """,""" & pkm1_iv3.Value & """,""" & _
                pkm1_iv4.Value & """,""" & pkm1_iv5.Value & """,""" & _
                pkm1_iv6.Value & """,""" & pkm1_ev1.Value & """,""" & _
                pkm1_ev2.Value & """,""" & pkm1_ev3.Value & """,""" & _
                pkm1_ev4.Value & """,""" & pkm1_ev5.Value & """,""" & _
                pkm1_ev6.Value & """,""" & pkm1_stat1.Text & """,""" & _
                pkm1_stat2.Text & """,""" & pkm1_stat3.Text & """,""" & _
                pkm1_stat4.Text & """,""" & pkm1_stat5.Text & """,""" & _
                pkm1_stat6.Text & """,""" & shinies(0) & """,""" & _
                pkm1_move1.Text & """,""" & pkm1_move2.Text & """,""" & _
                pkm1_move3.Text & """,""" & pkm1_move4.Text & """,""" & _
                pkm1_notes.Text & """" & vbCrLf & _
                """" & pkm2.Text & """,""" & pkm2_nickname.Text & """,""" & _
                pkm2_helditem.Text & """,""" & pkm2_gender.Text & """,""" & _
                pkm2_ability.Text & """,""" & pkm2_nature.Text & """,""" & _
                pkm2_level.Value & """,""" & pkm2_iv1.Value & """,""" & _
                pkm2_iv2.Value & """,""" & pkm2_iv3.Value & """,""" & _
                pkm2_iv4.Value & """,""" & pkm2_iv5.Value & """,""" & _
                pkm2_iv6.Value & """,""" & pkm2_ev1.Value & """,""" & _
                pkm2_ev2.Value & """,""" & pkm2_ev3.Value & """,""" & _
                pkm2_ev4.Value & """,""" & pkm2_ev5.Value & """,""" & _
                pkm2_ev6.Value & """,""" & pkm2_stat1.Text & """,""" & _
                pkm2_stat2.Text & """,""" & pkm2_stat3.Text & """,""" & _
                pkm2_stat4.Text & """,""" & pkm2_stat5.Text & """,""" & _
                pkm2_stat6.Text & """,""" & shinies(1) & """,""" & _
                pkm2_move1.Text & """,""" & pkm2_move2.Text & """,""" & _
                pkm2_move3.Text & """,""" & pkm2_move4.Text & """,""" & _
                pkm2_notes.Text & """" & vbCrLf & _
                """" & pkm3.Text & """,""" & pkm3_nickname.Text & """,""" & _
                pkm3_helditem.Text & """,""" & pkm3_gender.Text & """,""" & _
                pkm3_ability.Text & """,""" & pkm3_nature.Text & """,""" & _
                pkm3_level.Value & """,""" & pkm3_iv1.Value & """,""" & _
                pkm3_iv2.Value & """,""" & pkm3_iv3.Value & """,""" & _
                pkm3_iv4.Value & """,""" & pkm3_iv5.Value & """,""" & _
                pkm3_iv6.Value & """,""" & pkm3_ev1.Value & """,""" & _
                pkm3_ev2.Value & """,""" & pkm3_ev3.Value & """,""" & _
                pkm3_ev4.Value & """,""" & pkm3_ev5.Value & """,""" & _
                pkm3_ev6.Value & """,""" & pkm3_stat1.Text & """,""" & _
                pkm3_stat2.Text & """,""" & pkm3_stat3.Text & """,""" & _
                pkm3_stat4.Text & """,""" & pkm3_stat5.Text & """,""" & _
                pkm3_stat6.Text & """,""" & shinies(2) & """,""" & _
                pkm3_move1.Text & """,""" & pkm3_move2.Text & """,""" & _
                pkm3_move3.Text & """,""" & pkm3_move4.Text & """,""" & _
                pkm3_notes.Text & """" & vbCrLf & _
                """" & pkm4.Text & """,""" & pkm4_nickname.Text & """,""" & _
                pkm4_helditem.Text & """,""" & pkm4_gender.Text & """,""" & _
                pkm4_ability.Text & """,""" & pkm4_nature.Text & """,""" & _
                pkm4_level.Value & """,""" & pkm4_iv1.Value & """,""" & _
                pkm4_iv2.Value & """,""" & pkm4_iv3.Value & """,""" & _
                pkm4_iv4.Value & """,""" & pkm4_iv5.Value & """,""" & _
                pkm4_iv6.Value & """,""" & pkm4_ev1.Value & """,""" & _
                pkm4_ev2.Value & """,""" & pkm4_ev3.Value & """,""" & _
                pkm4_ev4.Value & """,""" & pkm4_ev5.Value & """,""" & _
                pkm4_ev6.Value & """,""" & pkm4_stat1.Text & """,""" & _
                pkm4_stat2.Text & """,""" & pkm4_stat3.Text & """,""" & _
                pkm4_stat4.Text & """,""" & pkm4_stat5.Text & """,""" & _
                pkm4_stat6.Text & """,""" & shinies(3) & """,""" & _
                pkm4_move1.Text & """,""" & pkm4_move2.Text & """,""" & _
                pkm4_move3.Text & """,""" & pkm4_move4.Text & """,""" & _
                pkm4_notes.Text & """" & vbCrLf & _
                """" & pkm5.Text & """,""" & pkm5_nickname.Text & """,""" & _
                pkm5_helditem.Text & """,""" & pkm5_gender.Text & """,""" & _
                pkm5_ability.Text & """,""" & pkm5_nature.Text & """,""" & _
                pkm5_level.Value & """,""" & pkm5_iv1.Value & """,""" & _
                pkm5_iv2.Value & """,""" & pkm5_iv3.Value & """,""" & _
                pkm5_iv4.Value & """,""" & pkm5_iv5.Value & """,""" & _
                pkm5_iv6.Value & """,""" & pkm5_ev1.Value & """,""" & _
                pkm5_ev2.Value & """,""" & pkm5_ev3.Value & """,""" & _
                pkm5_ev4.Value & """,""" & pkm5_ev5.Value & """,""" & _
                pkm5_ev6.Value & """,""" & pkm5_stat1.Text & """,""" & _
                pkm5_stat2.Text & """,""" & pkm5_stat3.Text & """,""" & _
                pkm5_stat4.Text & """,""" & pkm5_stat5.Text & """,""" & _
                pkm5_stat6.Text & """,""" & shinies(4) & """,""" & _
                pkm5_move1.Text & """,""" & pkm5_move2.Text & """,""" & _
                pkm5_move3.Text & """,""" & pkm5_move4.Text & """,""" & _
                pkm5_notes.Text & """" & vbCrLf & _
                """" & pkm6.Text & """,""" & pkm6_nickname.Text & """,""" & _
                pkm6_helditem.Text & """,""" & pkm6_gender.Text & """,""" & _
                pkm6_ability.Text & """,""" & pkm6_nature.Text & """,""" & _
                pkm6_level.Value & """,""" & pkm6_iv1.Value & """,""" & _
                pkm6_iv2.Value & """,""" & pkm6_iv3.Value & """,""" & _
                pkm6_iv4.Value & """,""" & pkm6_iv5.Value & """,""" & _
                pkm6_iv6.Value & """,""" & pkm6_ev1.Value & """,""" & _
                pkm6_ev2.Value & """,""" & pkm6_ev3.Value & """,""" & _
                pkm6_ev4.Value & """,""" & pkm6_ev5.Value & """,""" & _
                pkm6_ev6.Value & """,""" & pkm6_stat1.Text & """,""" & _
                pkm6_stat2.Text & """,""" & pkm6_stat3.Text & """,""" & _
                pkm6_stat4.Text & """,""" & pkm6_stat5.Text & """,""" & _
                pkm6_stat6.Text & """,""" & shinies(5) & """,""" & _
                pkm6_move1.Text & """,""" & pkm6_move2.Text & """,""" & _
                pkm6_move3.Text & """,""" & pkm6_move4.Text & """,""" & _
                pkm6_notes.Text & """"
            
            Dim sw As New StreamWriter(CSVFilename, False, System.Text.Encoding.UTF8)
            sw.Write(WriteString)
            sw.Close()

            Cursor = Cursors.Default
        Catch ex As Exception
            Cursor = Cursors.Default
            MsgBox("The team could not be exported:" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub ExportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportToolStripMenuItem.Click
        If SaveHTMLDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            DoSaveHTML(SaveHTMLDialog.FileName)
        End If
    End Sub

    'htmlspecialchars
    Public Function hsc(ByVal MyString As String) As String
        Return MyString.Replace("&", "&amp;").Replace("""", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;").Replace(vbLf, "<br />").Replace(vbCrLf, "<br />")
    End Function

    Public Function BuildHTML(ByRef PrintMode As Boolean) As String
        Dim shinies() As String = {"", "", "", "", "", ""}
        If pkm1_shiny.CheckState = CheckState.Checked Then
            shinies(0) = "<br /><span style=""color: red;"">Shiny</span>"
        End If
        If pkm2_shiny.CheckState = CheckState.Checked Then
            shinies(1) = "<br /><span style=""color: red;"">Shiny</span>"
        End If
        If pkm3_shiny.CheckState = CheckState.Checked Then
            shinies(2) = "<br /><span style=""color: red;"">Shiny</span>"
        End If
        If pkm4_shiny.CheckState = CheckState.Checked Then
            shinies(3) = "<br /><span style=""color: red;"">Shiny</span>"
        End If
        If pkm5_shiny.CheckState = CheckState.Checked Then
            shinies(4) = "<br /><span style=""color: red;"">Shiny</span>"
        End If
        If pkm6_shiny.CheckState = CheckState.Checked Then
            shinies(5) = "<br /><span style=""color: red;"">Shiny</span>"
        End If
        Dim H1 As String = ""
        Dim Title As String = "Exported Pok&eacute;mon Team"
        If prettyFilename.Length > 0 Then
            H1 = "<h1>" & hsc(prettyFilename) & "</h1>"
            Title &= ": " & hsc(prettyFilename)
        End If

        Dim Images() As String = {"", "", "", "", "", ""}
        Dim ms1 As New MemoryStream
        Dim ms2 As New MemoryStream
        Dim ms3 As New MemoryStream
        Dim ms4 As New MemoryStream
        Dim ms5 As New MemoryStream
        Dim ms6 As New MemoryStream
        pkm1_image.Image.Save(ms1, System.Drawing.Imaging.ImageFormat.Png)
        Images(0) = Convert.ToBase64String(ms1.ToArray())
        pkm2_image.Image.Save(ms2, System.Drawing.Imaging.ImageFormat.Png)
        Images(1) = Convert.ToBase64String(ms2.ToArray())
        pkm3_image.Image.Save(ms3, System.Drawing.Imaging.ImageFormat.Png)
        Images(2) = Convert.ToBase64String(ms3.ToArray())
        pkm4_image.Image.Save(ms4, System.Drawing.Imaging.ImageFormat.Png)
        Images(3) = Convert.ToBase64String(ms4.ToArray())
        pkm5_image.Image.Save(ms5, System.Drawing.Imaging.ImageFormat.Png)
        Images(4) = Convert.ToBase64String(ms5.ToArray())
        pkm6_image.Image.Save(ms6, System.Drawing.Imaging.ImageFormat.Png)
        Images(5) = Convert.ToBase64String(ms6.ToArray())
        ms1.Close()
        ms2.Close()
        ms3.Close()
        ms4.Close()
        ms5.Close()


        Dim PrintString As String = ""
        If PrintMode Then
            PrintString = " onload=""window.print();"""
        End If

        Return "<!DOCTYPE html>" & vbCrLf & _
                        "<html lang=""en"">" & vbCrLf & _
                        "<!-- Generated by Team Viewer by Cat" & vbCrLf & "     http://victoryroad.net/showthread.php?t=10904 -->" & _
                        vbCrLf & "<head><title>" & Title & "</title><meta charset=""utf-8"" />" & _
                        "<style type=""text/css"">td{font-size: 11pt; border: 1px #BBB inset; padding: 3px;}th{text-align: left;}</style></head><body" & PrintString & ">" & H1 & vbCrLf & _
                        "<table><tr>" & _
                        "<th>Pokémon</th><th style=""width: 115px;"">Information</th>" & _
                        "<th>HP</th><th>Atk</th><th>Def</th><th>SpA</th><th>SpD</th><th>Spd</th>" & _
                        "<th style=""width: 100px;"">Moves</th><th style=""width: 200px;"">Notes</th></tr>" & vbCrLf & _
                        "<tr><td><b>" & hsc(pkm1.Text) & "</b><br /><i>" & hsc(pkm1_nickname.Text) & "</i>" & shinies(0) & "</td><td>@ " & _
                        pkm1_helditem.Text & "<br />" & pkm1_gender.Text & "<br />" & _
                        pkm1_ability.Text & "<br />" & pkm1_nature.Text & "<br />Lv. " & _
                        pkm1_level.Value & "</td><td style=""text-align: right;"">" & _
                        pkm1_iv1.Value & "<br />" & pkm1_ev1.Value & "<br />" & pkm1_stat1.Text & "</td><td style=""text-align: right;"">" & _
                        pkm1_iv2.Value & "<br />" & pkm1_ev2.Value & "<br />" & pkm1_stat2.Text & "</td><td style=""text-align: right;"">" & _
                        pkm1_iv3.Value & "<br />" & pkm1_ev3.Value & "<br />" & pkm1_stat3.Text & "</td><td style=""text-align: right;"">" & _
                        pkm1_iv4.Value & "<br />" & pkm1_ev4.Value & "<br />" & pkm1_stat4.Text & "</td><td style=""text-align: right;"">" & _
                        pkm1_iv5.Value & "<br />" & pkm1_ev5.Value & "<br />" & pkm1_stat5.Text & "</td><td style=""text-align: right;"">" & _
                        pkm1_iv6.Value & "<br />" & pkm1_ev6.Value & "<br />" & pkm1_stat6.Text & "</td><td>" & _
                        pkm1_move1.Text & "<br />" & pkm1_move2.Text & "<br />" & _
                        pkm1_move3.Text & "<br />" & pkm1_move4.Text & "</td><td>" & _
                        hsc(pkm1_notes.Text) & "&nbsp;</td></tr>" & vbCrLf & _
                        "<tr><td><b>" & hsc(pkm2.Text) & "</b><br /><i>" & hsc(pkm2_nickname.Text) & "</i>" & shinies(1) & "</td><td>@ " & _
                        pkm2_helditem.Text & "<br />" & pkm2_gender.Text & "<br />" & _
                        pkm2_ability.Text & "<br />" & pkm2_nature.Text & "<br />Lv. " & _
                        pkm2_level.Value & "</td><td style=""text-align: right;"">" & _
                        pkm2_iv1.Value & "<br />" & pkm2_ev1.Value & "<br />" & pkm2_stat1.Text & "</td><td style=""text-align: right;"">" & _
                        pkm2_iv2.Value & "<br />" & pkm2_ev2.Value & "<br />" & pkm2_stat2.Text & "</td><td style=""text-align: right;"">" & _
                        pkm2_iv3.Value & "<br />" & pkm2_ev3.Value & "<br />" & pkm2_stat3.Text & "</td><td style=""text-align: right;"">" & _
                        pkm2_iv4.Value & "<br />" & pkm2_ev4.Value & "<br />" & pkm2_stat4.Text & "</td><td style=""text-align: right;"">" & _
                        pkm2_iv5.Value & "<br />" & pkm2_ev5.Value & "<br />" & pkm2_stat5.Text & "</td><td style=""text-align: right;"">" & _
                        pkm2_iv6.Value & "<br />" & pkm2_ev6.Value & "<br />" & pkm2_stat6.Text & "</td><td>" & _
                        pkm2_move1.Text & "<br />" & pkm2_move2.Text & "<br />" & _
                        pkm2_move3.Text & "<br />" & pkm2_move4.Text & "</td><td>" & _
                        hsc(pkm2_notes.Text) & "&nbsp;</td></tr>" & vbCrLf & _
                        "<tr><td><b>" & hsc(pkm3.Text) & "</b><br /><i>" & hsc(pkm3_nickname.Text) & "</i>" & shinies(2) & "</td><td>@ " & _
                        pkm3_helditem.Text & "<br />" & pkm3_gender.Text & "<br />" & _
                        pkm3_ability.Text & "<br />" & pkm3_nature.Text & "<br />Lv. " & _
                        pkm3_level.Value & "</td><td style=""text-align: right;"">" & _
                        pkm3_iv1.Value & "<br />" & pkm3_ev1.Value & "<br />" & pkm3_stat1.Text & "</td><td style=""text-align: right;"">" & _
                        pkm3_iv2.Value & "<br />" & pkm3_ev2.Value & "<br />" & pkm3_stat2.Text & "</td><td style=""text-align: right;"">" & _
                        pkm3_iv3.Value & "<br />" & pkm3_ev3.Value & "<br />" & pkm3_stat3.Text & "</td><td style=""text-align: right;"">" & _
                        pkm3_iv4.Value & "<br />" & pkm3_ev4.Value & "<br />" & pkm3_stat4.Text & "</td><td style=""text-align: right;"">" & _
                        pkm3_iv5.Value & "<br />" & pkm3_ev5.Value & "<br />" & pkm3_stat5.Text & "</td><td style=""text-align: right;"">" & _
                        pkm3_iv6.Value & "<br />" & pkm3_ev6.Value & "<br />" & pkm3_stat6.Text & "</td><td>" & _
                        pkm3_move1.Text & "<br />" & pkm3_move2.Text & "<br />" & _
                        pkm3_move3.Text & "<br />" & pkm3_move4.Text & "</td><td>" & _
                        hsc(pkm3_notes.Text) & "&nbsp;</td></tr>" & vbCrLf & _
                        "<tr><td><b>" & hsc(pkm4.Text) & "</b><br /><i>" & hsc(pkm4_nickname.Text) & "</i>" & shinies(3) & "</td><td>@ " & _
                        pkm4_helditem.Text & "<br />" & pkm4_gender.Text & "<br />" & _
                        pkm4_ability.Text & "<br />" & pkm4_nature.Text & "<br />Lv. " & _
                        pkm4_level.Value & "</td><td style=""text-align: right;"">" & _
                        pkm4_iv1.Value & "<br />" & pkm4_ev1.Value & "<br />" & pkm4_stat1.Text & "</td><td style=""text-align: right;"">" & _
                        pkm4_iv2.Value & "<br />" & pkm4_ev2.Value & "<br />" & pkm4_stat2.Text & "</td><td style=""text-align: right;"">" & _
                        pkm4_iv3.Value & "<br />" & pkm4_ev3.Value & "<br />" & pkm4_stat3.Text & "</td><td style=""text-align: right;"">" & _
                        pkm4_iv4.Value & "<br />" & pkm4_ev4.Value & "<br />" & pkm4_stat4.Text & "</td><td style=""text-align: right;"">" & _
                        pkm4_iv5.Value & "<br />" & pkm4_ev5.Value & "<br />" & pkm4_stat5.Text & "</td><td style=""text-align: right;"">" & _
                        pkm4_iv6.Value & "<br />" & pkm4_ev6.Value & "<br />" & pkm4_stat6.Text & "</td><td>" & _
                        pkm4_move1.Text & "<br />" & pkm4_move2.Text & "<br />" & _
                        pkm4_move3.Text & "<br />" & pkm4_move4.Text & "</td><td>" & _
                        hsc(pkm4_notes.Text) & "&nbsp;</td></tr>" & vbCrLf & _
                        "<tr><td><b>" & hsc(pkm5.Text) & "</b><br /><i>" & hsc(pkm5_nickname.Text) & "</i>" & shinies(4) & "</td><td>@ " & _
                        pkm5_helditem.Text & "<br />" & pkm5_gender.Text & "<br />" & _
                        pkm5_ability.Text & "<br />" & pkm5_nature.Text & "<br />Lv. " & _
                        pkm5_level.Value & "</td><td style=""text-align: right;"">" & _
                        pkm5_iv1.Value & "<br />" & pkm5_ev1.Value & "<br />" & pkm5_stat1.Text & "</td><td style=""text-align: right;"">" & _
                        pkm5_iv2.Value & "<br />" & pkm5_ev2.Value & "<br />" & pkm5_stat2.Text & "</td><td style=""text-align: right;"">" & _
                        pkm5_iv3.Value & "<br />" & pkm5_ev3.Value & "<br />" & pkm5_stat3.Text & "</td><td style=""text-align: right;"">" & _
                        pkm5_iv4.Value & "<br />" & pkm5_ev4.Value & "<br />" & pkm5_stat4.Text & "</td><td style=""text-align: right;"">" & _
                        pkm5_iv5.Value & "<br />" & pkm5_ev5.Value & "<br />" & pkm5_stat5.Text & "</td><td style=""text-align: right;"">" & _
                        pkm5_iv6.Value & "<br />" & pkm5_ev6.Value & "<br />" & pkm5_stat6.Text & "</td><td>" & _
                        pkm5_move1.Text & "<br />" & pkm5_move2.Text & "<br />" & _
                        pkm5_move3.Text & "<br />" & pkm5_move4.Text & "</td><td>" & _
                        hsc(pkm5_notes.Text) & "&nbsp;</td></tr>" & vbCrLf & _
                        "<tr><td><b>" & hsc(pkm6.Text) & "</b><br /><i>" & hsc(pkm6_nickname.Text) & "</i>" & shinies(5) & "</td><td>@ " & _
                        pkm6_helditem.Text & "<br />" & pkm6_gender.Text & "<br />" & _
                        pkm6_ability.Text & "<br />" & pkm6_nature.Text & "<br />Lv. " & _
                        pkm6_level.Value & "</td><td style=""text-align: right;"">" & _
                        pkm6_iv1.Value & "<br />" & pkm6_ev1.Value & "<br />" & pkm6_stat1.Text & "</td><td style=""text-align: right;"">" & _
                        pkm6_iv2.Value & "<br />" & pkm6_ev2.Value & "<br />" & pkm6_stat2.Text & "</td><td style=""text-align: right;"">" & _
                        pkm6_iv3.Value & "<br />" & pkm6_ev3.Value & "<br />" & pkm6_stat3.Text & "</td><td style=""text-align: right;"">" & _
                        pkm6_iv4.Value & "<br />" & pkm6_ev4.Value & "<br />" & pkm6_stat4.Text & "</td><td style=""text-align: right;"">" & _
                        pkm6_iv5.Value & "<br />" & pkm6_ev5.Value & "<br />" & pkm6_stat5.Text & "</td><td style=""text-align: right;"">" & _
                        pkm6_iv6.Value & "<br />" & pkm6_ev6.Value & "<br />" & pkm6_stat6.Text & "</td><td>" & _
                        pkm6_move1.Text & "<br />" & pkm6_move2.Text & "<br />" & _
                        pkm6_move3.Text & "<br />" & pkm6_move4.Text & "</td><td>" & _
                        hsc(pkm6_notes.Text) & "&nbsp;</td></tr>" & vbCrLf & _
                        "</table><br />" & _
                        "<img src=""data:image/png;base64," & Images(0) & """ alt=""" & hsc(pkm1.Text) & """ /> " & _
                        "<img src=""data:image/png;base64," & Images(1) & """ alt=""" & hsc(pkm2.Text) & """ /> " & _
                        "<img src=""data:image/png;base64," & Images(2) & """ alt=""" & hsc(pkm3.Text) & """ /> " & _
                        "<img src=""data:image/png;base64," & Images(3) & """ alt=""" & hsc(pkm4.Text) & """ /> " & _
                        "<img src=""data:image/png;base64," & Images(4) & """ alt=""" & hsc(pkm5.Text) & """ /> " & _
                        "<img src=""data:image/png;base64," & Images(5) & """ alt=""" & hsc(pkm6.Text) & """ /> " & _
                        "</body></html>"
    End Function

    Public Sub DoSaveHTML(ByRef HTMLFilename As String)
        Try
            Cursor = Cursors.WaitCursor
            Dim WriteString As String = BuildHTML(False)

            Dim sw As New StreamWriter(HTMLFilename, False, System.Text.Encoding.UTF8)
            sw.Write(WriteString)
            sw.Close()

            Cursor = Cursors.Default
        Catch ex As Exception
            Cursor = Cursors.Default
            MsgBox("The team could not be exported:" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub


    Private Sub PrintToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Dim PrintBrowser As New WebBrowser
        PrintBrowser.DocumentText = " "
        PrintBrowser.Document.Write(BuildHTML(False))
        PrintBrowser.ShowPrintDialog()
        PrintBrowser.Dispose()
    End Sub

    Private Sub ExportTextToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportTextToolStripMenuItem.Click
        Dim PlainString As String = ""
        If pkm1.SelectedIndex > 0 Then
            If pkm1_nickname.Text.Length > 0 Then
                PlainString &= pokemonNames(0) & " (" & pkm1_nickname.Text & ")"
            Else
                PlainString &= pokemonNames(0)
            End If
            If pkm1_helditem.SelectedIndex > 0 Then
                PlainString &= " @ " & pkm1_helditem.Text
            Else
                PlainString &= " @ No item"
            End If
            PlainString &= vbCrLf & pkm1_ability.Text
            If pkm1_nature.SelectedIndex > 0 Then
                PlainString &= "/" & pkm1_nature.Text
            End If
            If pkm1_move1.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm1_move1.Text
            End If
            If pkm1_move2.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm1_move2.Text
            End If
            If pkm1_move3.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm1_move3.Text
            End If
            If pkm1_move4.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm1_move4.Text
            End If
        End If
        If pkm2.SelectedIndex > 0 Then
            If PlainString.Length > 0 Then
                PlainString &= vbCrLf & vbCrLf
            End If
            If pkm2_nickname.Text.Length > 0 Then
                PlainString &= pokemonNames(1) & " (" & pkm2_nickname.Text & ")"
            Else
                PlainString &= pokemonNames(1)
            End If
            If pkm2_helditem.SelectedIndex > 0 Then
                PlainString &= " @ " & pkm2_helditem.Text
            Else
                PlainString &= " @ No item"
            End If
            PlainString &= vbCrLf & pkm2_ability.Text
            If pkm2_nature.SelectedIndex > 0 Then
                PlainString &= "/" & pkm2_nature.Text
            End If
            If pkm2_move1.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm2_move1.Text
            End If
            If pkm2_move2.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm2_move2.Text
            End If
            If pkm2_move3.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm2_move3.Text
            End If
            If pkm2_move4.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm2_move4.Text
            End If
        End If
        If pkm3.SelectedIndex > 0 Then
            If PlainString.Length > 0 Then
                PlainString &= vbCrLf & vbCrLf
            End If
            If pkm3_nickname.Text.Length > 0 Then
                PlainString &= pokemonNames(2) & " (" & pkm3_nickname.Text & ")"
            Else
                PlainString &= pokemonNames(2)
            End If
            If pkm3_helditem.SelectedIndex > 0 Then
                PlainString &= " @ " & pkm3_helditem.Text
            Else
                PlainString &= " @ No item"
            End If
            PlainString &= vbCrLf & pkm3_ability.Text
            If pkm3_nature.SelectedIndex > 0 Then
                PlainString &= "/" & pkm3_nature.Text
            End If
            If pkm3_move1.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm3_move1.Text
            End If
            If pkm3_move2.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm3_move2.Text
            End If
            If pkm3_move3.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm3_move3.Text
            End If
            If pkm3_move4.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm3_move4.Text
            End If
        End If
        If pkm4.SelectedIndex > 0 Then
            If PlainString.Length > 0 Then
                PlainString &= vbCrLf & vbCrLf
            End If
            If pkm4_nickname.Text.Length > 0 Then
                PlainString &= pokemonNames(3) & " (" & pkm4_nickname.Text & ")"
            Else
                PlainString &= pokemonNames(3)
            End If
            If pkm4_helditem.SelectedIndex > 0 Then
                PlainString &= " @ " & pkm4_helditem.Text
            Else
                PlainString &= " @ No item"
            End If
            PlainString &= vbCrLf & pkm4_ability.Text
            If pkm4_nature.SelectedIndex > 0 Then
                PlainString &= "/" & pkm4_nature.Text
            End If
            If pkm4_move1.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm4_move1.Text
            End If
            If pkm4_move2.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm4_move2.Text
            End If
            If pkm4_move3.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm4_move3.Text
            End If
            If pkm4_move4.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm4_move4.Text
            End If
        End If
        If pkm5.SelectedIndex > 0 Then
            If PlainString.Length > 0 Then
                PlainString &= vbCrLf & vbCrLf
            End If
            If pkm5_nickname.Text.Length > 0 Then
                PlainString &= pokemonNames(4) & " (" & pkm5_nickname.Text & ")"
            Else
                PlainString &= pokemonNames(4)
            End If
            If pkm5_helditem.SelectedIndex > 0 Then
                PlainString &= " @ " & pkm5_helditem.Text
            Else
                PlainString &= " @ No item"
            End If
            PlainString &= vbCrLf & pkm5_ability.Text
            If pkm5_nature.SelectedIndex > 0 Then
                PlainString &= "/" & pkm5_nature.Text
            End If
            If pkm5_move1.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm5_move1.Text
            End If
            If pkm5_move2.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm5_move2.Text
            End If
            If pkm5_move3.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm5_move3.Text
            End If
            If pkm5_move4.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm5_move4.Text
            End If
        End If
        If pkm6.SelectedIndex > 0 Then
            If PlainString.Length > 0 Then
                PlainString &= vbCrLf & vbCrLf
            End If
            If pkm6_nickname.Text.Length > 0 Then
                PlainString &= pokemonNames(5) & " (" & pkm6_nickname.Text & ")"
            Else
                PlainString &= pokemonNames(5)
            End If
            If pkm6_helditem.SelectedIndex > 0 Then
                PlainString &= " @ " & pkm6_helditem.Text
            Else
                PlainString &= " @ No item"
            End If
            PlainString &= vbCrLf & pkm6_ability.Text
            If pkm6_nature.SelectedIndex > 0 Then
                PlainString &= "/" & pkm6_nature.Text
            End If
            If pkm6_move1.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm6_move1.Text
            End If
            If pkm6_move2.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm6_move2.Text
            End If
            If pkm6_move3.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm6_move3.Text
            End If
            If pkm6_move4.SelectedIndex > 0 Then
                PlainString &= vbCrLf & "- " & pkm6_move4.Text
            End If
        End If
        TextExport.TextExportString.Text = PlainString
        TextExport.Show()
    End Sub
End Class
