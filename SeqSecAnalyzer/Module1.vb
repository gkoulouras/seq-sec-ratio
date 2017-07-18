Imports OfficeOpenXml
Imports System.Drawing
Imports System.IO

Module Module1

    Public SeqDTable As DataTable
    Public SecDTable As DataTable
    Public sum As Integer
    Sub Main()

        Dim pdb_seq_name As String = String.Empty
        Dim pdb_sec_name As String = String.Empty
        Dim dIndex As Integer = 0
        Dim source As String = String.Empty
        Dim stpw As New Stopwatch
        'Dim path As String = "C:\Users\grigo\Dropbox\PROJECTS\PEZ_THESIS\ss\ss2.txt"
        Dim linecounter As Integer = 0
        Dim FastaType As String = String.Empty

        Try


            Console.WriteLine("The application started..." & vbCrLf)
            Console.WriteLine("Please declare the full path where the ss.txt file is located in:" & vbCrLf)

            Dim path As String = Console.ReadLine
            stpw.Restart()
            Dim sr As New IO.StreamReader(path)

            Console.WriteLine("The following rows of the file have been retrieved:" & vbCrLf)

            'Create a datatable to store the results
            SeqDTable = CreateDataTable4Seq()
            SecDTable = CreateDataTable4Sec()

            Dim s As String = String.Empty
            s = sr.ReadLine
            linecounter = linecounter + 1

            Do While Not sr.EndOfStream

                If s.Contains(">") And s.Contains("sequence") Then 'sequence
                    FastaType = "sequence"
                    dIndex = s.IndexOf(":")
                    pdb_seq_name = s.Substring(1, dIndex - 1)
                ElseIf s.Contains(">") And s.Contains("secstr") Then 'secondary structure
                    FastaType = "secondary"
                    dIndex = s.IndexOf(":")
                    pdb_sec_name = s.Substring(1, dIndex - 1)
                End If

                Select Case FastaType
                    Case "sequence"
                        If s.Contains(":A:") Then   'read the first isoform
                            Dim seqstring As String = String.Empty
                            s = sr.ReadLine()
                            linecounter = linecounter + 1
                            While Not s.Contains(">")
                                seqstring = seqstring + s
                                s = sr.ReadLine()
                                linecounter = linecounter + 1
                                If s Is Nothing Then
                                    EstimateAminoAcidQuantities(pdb_seq_name, seqstring)
                                    GoTo EndLoop
                                End If
                            End While
                            EstimateAminoAcidQuantities(pdb_seq_name, seqstring)
                        Else                        'skip the other isoforms
                            s = sr.ReadLine()
                            linecounter = linecounter + 1
                            While Not s.Contains(">")
                                s = sr.ReadLine()
                                linecounter = linecounter + 1
                                If s Is Nothing Then
                                    GoTo EndLoop
                                End If
                            End While
                        End If
                    Case "secondary"
                        If s.Contains(":A:") Then   'read the first isoform
                            Dim secstring As String = String.Empty
                            s = sr.ReadLine()
                            linecounter = linecounter + 1
                            While Not s.Contains(">")
                                secstring = secstring + s
                                s = sr.ReadLine()
                                linecounter = linecounter + 1
                                If s Is Nothing Then
                                    SaveSecondaryStructures(pdb_sec_name, secstring)
                                    GoTo EndLoop
                                End If
                            End While
                            SaveSecondaryStructures(pdb_sec_name, secstring)
                        Else                        'skip the other isoforms
                            s = sr.ReadLine()
                            linecounter = linecounter + 1
                            While Not s.Contains(">")
                                s = sr.ReadLine()
                                linecounter = linecounter + 1
                                If s Is Nothing Then
                                    GoTo EndLoop
                                End If
                            End While
                        End If
                End Select

                Console.SetCursorPosition(0, Console.CursorTop)
                Console.Write(linecounter.ToString)

            Loop

EndLoop:    Console.SetCursorPosition(0, Console.CursorTop)
            Console.Write(linecounter.ToString)
            sr.Close()

            CalculatePercentagesAndPrintResults()

        Catch e As Exception
            Console.WriteLine("An error occured.")
            Console.WriteLine("The error message is: " & e.Message)
            Console.WriteLine(vbCrLf & "Press enter to exit...")
            Console.ReadLine()
            Environment.Exit(0)
        End Try
        stpw.Stop()
        Console.WriteLine(vbCrLf & vbCrLf & "The application was terminated.")
        Console.WriteLine("Elapsed time: " & stpw.Elapsed.ToString)
        Console.WriteLine("Press enter to exit...")
        Console.ReadLine()
        Environment.Exit(0)
    End Sub

    Private Function CreateDataTable4Seq() 'is used as constructor to store the output results
        Dim table As New DataTable
        table.Columns.Add("pdb_ID", GetType(String))
        table.Columns.Add("fasta", GetType(String))
        table.Columns.Add("A", GetType(String))
        table.Columns.Add("B", GetType(String))
        table.Columns.Add("C", GetType(String))
        table.Columns.Add("D", GetType(String))
        table.Columns.Add("E", GetType(String))
        table.Columns.Add("F", GetType(String))
        table.Columns.Add("G", GetType(String))
        table.Columns.Add("H", GetType(String))
        table.Columns.Add("I", GetType(String))
        table.Columns.Add("J", GetType(String))
        table.Columns.Add("K", GetType(String))
        table.Columns.Add("L", GetType(String))
        table.Columns.Add("M", GetType(String))
        table.Columns.Add("N", GetType(String))
        table.Columns.Add("O", GetType(String))
        table.Columns.Add("P", GetType(String))
        table.Columns.Add("Q", GetType(String))
        table.Columns.Add("R", GetType(String))
        table.Columns.Add("S", GetType(String))
        table.Columns.Add("T", GetType(String))
        table.Columns.Add("U", GetType(String))
        table.Columns.Add("V", GetType(String))
        table.Columns.Add("W", GetType(String))
        table.Columns.Add("X", GetType(String))
        table.Columns.Add("Y", GetType(String))
        table.Columns.Add("Z", GetType(String))
        Return table
    End Function

    Private Function CreateDataTable4Sec() 'is used as constructor to store the output results
        Dim table As New DataTable
        table.Columns.Add("pdb_ID", GetType(String))
        table.Columns.Add("sec_struct", GetType(String))
        Return table
    End Function

    Private Sub EstimateAminoAcidQuantities(ByVal pdb_identifier As String, ByVal FASTA As String)

        Dim A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z As Integer

        For Each charc As Char In FASTA
            Select Case charc
                Case "A" 'Alanine
                    A += 1
                Case "B" 'Asn or Asp
                    B += 1
                Case "C" 'Cysteine
                    C += 1
                Case "D" 'Aspartic acid
                    D += 1
                Case "E" 'Glutamic acid
                    E += 1
                Case "F" 'Phenylalanine
                    F += 1
                Case "G" 'Glycine
                    G += 1
                Case "H" 'Histidine
                    H += 1
                Case "I" 'Isoleucine
                    I += 1
                Case "J" 'Leucine (L) or Isoleucine (I)
                    J += 1
                Case "K" 'Lysine
                    K += 1
                Case "L" 'Leucine
                    L += 1
                Case "M" 'Methionine
                    M += 1
                Case "N" 'Asparagine
                    N += 1
                Case "O" 'Pyrrolysine
                    O += 1
                Case "P" 'Proline
                    P += 1
                Case "Q" 'Glutamine
                    Q += 1
                Case "R" 'Arginine
                    R += 1
                Case "S" 'Serine
                    S += 1
                Case "T" 'Threonine
                    T += 1
                Case "U" 'Selenocysteine
                    U += 1
                Case "V" 'Valine
                    V += 1
                Case "W" 'Tryptophan
                    W += 1
                Case "X" 'Any
                    X += 1
                Case "Y" 'Tyrosine
                    Y += 1
                Case "Z" 'Gln or Glu
                    Z += 1
                Case Else
                    Console.WriteLine(pdb_identifier & " - Invalid character found: " & charc.ToString)
            End Select
        Next
        'Fill the Datatable
        SeqDTable.Rows.Add(pdb_identifier, FASTA, A.ToString, B.ToString, C.ToString, D.ToString, E.ToString, F.ToString, G.ToString, H.ToString, I.ToString, J.ToString, K.ToString, L.ToString, M.ToString, N.ToString, O.ToString, P.ToString, Q.ToString, R.ToString, S.ToString, T.ToString, U.ToString, V.ToString, W.ToString, X.ToString, Y.ToString, Z.ToString)
    End Sub

    Private Sub SaveSecondaryStructures(ByVal pdb_identifier As String, ByVal secondary As String)
        SecDTable.Rows.Add(pdb_identifier, secondary)
    End Sub


    Private Sub CalculatePercentagesAndPrintResults()

        Dim path As String = Directory.GetCurrentDirectory()
        Dim pck As New ExcelPackage
        Dim ws, ws2, ws3 As ExcelWorksheet


        Dim generalcounter, Acounter, Bcounter, Ccounter, Dcounter, Ecounter, Fcounter, Gcounter, Hcounter, Icounter,
        Jcounter, Kcounter, Lcounter, Mcounter, Ncounter, Ocounter, Pcounter, Qcounter, Rcounter, Scounter,
        Tcounter, Ucounter, Vcounter, Wcounter, Xcounter, Ycounter, Zcounter As Integer

        Dim H_A, H_C, H_D, H_E, H_F, H_G, H_H, H_I, H_K, H_L, H_M, H_N, H_P, H_Q, H_R, H_S, H_T, H_V, H_W, H_Y,
        B_A, B_C, B_D, B_E, B_F, B_G, B_H, B_I, B_K, B_L, B_M, B_N, B_P, B_Q, B_R, B_S, B_T, B_V, B_W, B_Y,
        E_A, E_C, E_D, E_E, E_F, E_G, E_H, E_I, E_K, E_L, E_M, E_N, E_P, E_Q, E_R, E_S, E_T, E_V, E_W, E_Y,
        G_A, G_C, G_D, G_E, G_F, G_G, G_H, G_I, G_K, G_L, G_M, G_N, G_P, G_Q, G_R, G_S, G_T, G_V, G_W, G_Y,
        I_A, I_C, I_D, I_E, I_F, I_G, I_H, I_I, I_K, I_L, I_M, I_N, I_P, I_Q, I_R, I_S, I_T, I_V, I_W, I_Y,
        T_A, T_C, T_D, T_E, T_F, T_G, T_H, T_I, T_K, T_L, T_M, T_N, T_P, T_Q, T_R, T_S, T_T, T_V, T_W, T_Y,
        S_A, S_C, S_D, S_E, S_F, S_G, S_H, S_I, S_K, S_L, S_M, S_N, S_P, S_Q, S_R, S_S, S_T, S_V, S_W, S_Y As Integer


        'Calculate Percentages for Sequences
        For Each datarow As DataRow In SeqDTable.Rows
            For Each datacolumn As DataColumn In SeqDTable.Columns
                Select Case datacolumn.ColumnName
                    Case "A"
                        Acounter = Acounter + Convert.ToInt32(datarow("A"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("A"))
                    Case "B"
                        Bcounter = Bcounter + Convert.ToInt32(datarow("B"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("B"))
                    Case "C"
                        Ccounter = Ccounter + Convert.ToInt32(datarow("C"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("C"))
                    Case "D"
                        Dcounter = Dcounter + Convert.ToInt32(datarow("D"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("D"))
                    Case "E"
                        Ecounter = Ecounter + Convert.ToInt32(datarow("E"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("E"))
                    Case "F"
                        Fcounter = Fcounter + Convert.ToInt32(datarow("F"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("F"))
                    Case "G"
                        Gcounter = Gcounter + Convert.ToInt32(datarow("G"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("G"))
                    Case "H"
                        Hcounter = Hcounter + Convert.ToInt32(datarow("H"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("H"))
                    Case "I"
                        Icounter = Icounter + Convert.ToInt32(datarow("I"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("I"))
                    Case "J"
                        Jcounter = Jcounter + Convert.ToInt32(datarow("J"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("J"))
                    Case "K"
                        Kcounter = Kcounter + Convert.ToInt32(datarow("K"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("K"))
                    Case "L"
                        Lcounter = Lcounter + Convert.ToInt32(datarow("L"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("L"))
                    Case "M"
                        Mcounter = Mcounter + Convert.ToInt32(datarow("M"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("M"))
                    Case "N"
                        Ncounter = Ncounter + Convert.ToInt32(datarow("N"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("N"))
                    Case "O"
                        Ocounter = Ocounter + Convert.ToInt32(datarow("O"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("O"))
                    Case "P"
                        Pcounter = Pcounter + Convert.ToInt32(datarow("P"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("P"))
                    Case "Q"
                        Qcounter = Qcounter + Convert.ToInt32(datarow("Q"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("Q"))
                    Case "R"
                        Rcounter = Rcounter + Convert.ToInt32(datarow("R"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("R"))
                    Case "S"
                        Scounter = Scounter + Convert.ToInt32(datarow("S"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("S"))
                    Case "T"
                        Tcounter = Tcounter + Convert.ToInt32(datarow("T"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("T"))
                    Case "U"
                        Ucounter = Ucounter + Convert.ToInt32(datarow("U"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("U"))
                    Case "V"
                        Vcounter = Vcounter + Convert.ToInt32(datarow("V"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("V"))
                    Case "W"
                        Wcounter = Wcounter + Convert.ToInt32(datarow("W"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("W"))
                    Case "X"
                        Xcounter = Xcounter + Convert.ToInt32(datarow("X"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("X"))
                    Case "Y"
                        Ycounter = Ycounter + Convert.ToInt32(datarow("Y"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("Y"))
                    Case "Z"
                        Zcounter = Zcounter + Convert.ToInt32(datarow("Z"))
                        generalcounter = generalcounter + Convert.ToInt32(datarow("Z"))
                End Select
            Next
        Next

        Dim sumcounter As Integer = Acounter + Ccounter + Dcounter + Ecounter + Fcounter + Gcounter +
        Hcounter + Icounter + Kcounter + Lcounter + Mcounter + Ncounter + Pcounter + Qcounter +
        Rcounter + Scounter + Tcounter + Vcounter + Wcounter + Ycounter
        Dim skippedcounter As Integer = Bcounter + Jcounter + Ocounter + Ucounter + Xcounter + Zcounter


        ''Calculate Sequences - Secondary Structure Ratio
        If SeqDTable.Rows.Count <> SecDTable.Rows.Count Then
            Console.WriteLine(vbCrLf & "There is a mismatch between the sequences and the number of the secondary structures.")
            Console.WriteLine(SeqDTable.Rows.Count & " sequences found instead of " & SecDTable.Rows.Count & " secondary structures. Check your input file.")
            Return
        Else
            Console.WriteLine(vbCrLf & vbCrLf & "Please wait. The protein sequences are now being processed:" & vbCrLf)
            For i As Integer = 0 To SeqDTable.Rows.Count - 1
                If (SeqDTable.Rows(i).Item(0).ToString = SecDTable.Rows(i).Item(0).ToString) And (SeqDTable.Rows(i).Item(1).ToString.Length = SecDTable.Rows(i).Item(1).ToString.Length) Then
                    'Console.WriteLine(vbCrLf & SecDTable.Rows(i).Item(0).ToString & vbCrLf)
                    Console.SetCursorPosition(0, Console.CursorTop)
                    Console.Write(i.ToString)
                    Dim sequence As String = SeqDTable.Rows(i).Item(1).ToString
                    Dim sec_struct As String = SecDTable.Rows(i).Item(1).ToString
                    For j As Integer = 0 To SeqDTable.Rows(i).Item(1).ToString.Length - 1
                        'Console.WriteLine(sequence(j).ToString & " - " & sec_struct(j).ToString)
                        Select Case sec_struct(j).ToString
                            Case "H"    'alpha helix
                                Select Case sequence(j).ToString
                                    Case "A"    'Alanine
                                        H_A = H_A + 1
                                    Case "C"    'Cysteine
                                        H_C = H_C + 1
                                    Case "D"    'Aspartate
                                        H_D = H_D + 1
                                    Case "E"    'Glutamate
                                        H_E = H_E + 1
                                    Case "F"    'Phenylalanine
                                        H_F = H_F + 1
                                    Case "G"    'Glycine
                                        H_G = H_G + 1
                                    Case "H"    'Histidine
                                        H_H = H_H + 1
                                    Case "I"    'Isoleucine
                                        H_I = H_I + 1
                                    Case "K"    'Lysine
                                        H_K = H_K + 1
                                    Case "L"    'Leucine
                                        H_L = H_L + 1
                                    Case "M"    'Methionine
                                        H_M = H_M + 1
                                    Case "N"    'Asparagine
                                        H_N = H_N + 1
                                    Case "P"    'Proline
                                        H_P = H_P + 1
                                    Case "Q"    'Glutamine
                                        H_Q = H_Q + 1
                                    Case "R"    'Arginine
                                        H_R = H_R + 1
                                    Case "S"    'Serine
                                        H_S = H_S + 1
                                    Case "T"    'Threonine
                                        H_T = H_T + 1
                                    Case "V"    'Valine
                                        H_V = H_V + 1
                                    Case "W"    'Tryptophan
                                        H_W = H_W + 1
                                    Case "Y"    'Tyrosine
                                        H_Y = H_Y + 1
                                End Select
                            Case "B"    'residue in isolated beta-bridge
                                Select Case sequence(j).ToString
                                    Case "A"    'Alanine
                                        B_A = B_A + 1
                                    Case "C"    'Cysteine
                                        B_C = B_C + 1
                                    Case "D"    'Aspartate
                                        B_D = B_D + 1
                                    Case "E"    'Glutamate
                                        B_E = B_E + 1
                                    Case "F"    'Phenylalanine
                                        B_F = B_F + 1
                                    Case "G"    'Glycine
                                        B_G = B_G + 1
                                    Case "H"    'Histidine
                                        B_H = B_H + 1
                                    Case "I"    'Isoleucine
                                        B_I = B_I + 1
                                    Case "K"    'Lysine
                                        B_K = B_K + 1
                                    Case "L"    'Leucine
                                        B_L = B_L + 1
                                    Case "M"    'Methionine
                                        B_M = B_M + 1
                                    Case "N"    'Asparagine
                                        B_N = B_N + 1
                                    Case "P"    'Proline
                                        B_P = B_P + 1
                                    Case "Q"    'Glutamine
                                        B_Q = B_Q + 1
                                    Case "R"    'Arginine
                                        B_R = B_R + 1
                                    Case "S"    'Serine
                                        B_S = B_S + 1
                                    Case "T"    'Threonine
                                        B_T = B_T + 1
                                    Case "V"    'Valine
                                        B_V = B_V + 1
                                    Case "W"    'Tryptophan
                                        B_W = B_W + 1
                                    Case "Y"    'Tyrosine
                                        B_Y = B_Y + 1
                                End Select
                            Case "E"    'extended strand, participates in beta ladder
                                Select Case sequence(j).ToString
                                    Case "A"    'Alanine
                                        E_A = E_A + 1
                                    Case "C"    'Cysteine
                                        E_C = E_C + 1
                                    Case "D"    'Aspartate
                                        E_D = E_D + 1
                                    Case "E"    'Glutamate
                                        E_E = E_E + 1
                                    Case "F"    'Phenylalanine
                                        E_F = E_F + 1
                                    Case "G"    'Glycine
                                        E_G = E_G + 1
                                    Case "H"    'Histidine
                                        E_H = E_H + 1
                                    Case "I"    'Isoleucine
                                        E_I = E_I + 1
                                    Case "K"    'Lysine
                                        E_K = E_K + 1
                                    Case "L"    'Leucine
                                        E_L = E_L + 1
                                    Case "M"    'Methionine
                                        E_M = E_M + 1
                                    Case "N"    'Asparagine
                                        E_N = E_N + 1
                                    Case "P"    'Proline
                                        E_P = E_P + 1
                                    Case "Q"    'Glutamine
                                        E_Q = E_Q + 1
                                    Case "R"    'Arginine
                                        E_R = E_R + 1
                                    Case "S"    'Serine
                                        E_S = E_S + 1
                                    Case "T"    'Threonine
                                        E_T = E_T + 1
                                    Case "V"    'Valine
                                        E_V = E_V + 1
                                    Case "W"    'Tryptophan
                                        E_W = E_W + 1
                                    Case "Y"    'Tyrosine
                                        E_Y = E_Y + 1
                                End Select
                            Case "G"    '3-helix (3/10 helix)
                                Select Case sequence(j).ToString
                                    Case "A"    'Alanine
                                        G_A = G_A + 1
                                    Case "C"    'Cysteine
                                        G_C = G_C + 1
                                    Case "D"    'Aspartate
                                        G_D = G_D + 1
                                    Case "E"    'Glutamate
                                        G_E = G_E + 1
                                    Case "F"    'Phenylalanine
                                        G_F = G_F + 1
                                    Case "G"    'Glycine
                                        G_G = G_G + 1
                                    Case "H"    'Histidine
                                        G_H = G_H + 1
                                    Case "I"    'Isoleucine
                                        G_I = G_I + 1
                                    Case "K"    'Lysine
                                        G_K = G_K + 1
                                    Case "L"    'Leucine
                                        G_L = G_L + 1
                                    Case "M"    'Methionine
                                        G_M = G_M + 1
                                    Case "N"    'Asparagine
                                        G_N = G_N + 1
                                    Case "P"    'Proline
                                        G_P = G_P + 1
                                    Case "Q"    'Glutamine
                                        G_Q = G_Q + 1
                                    Case "R"    'Arginine
                                        G_R = G_R + 1
                                    Case "S"    'Serine
                                        G_S = G_S + 1
                                    Case "T"    'Threonine
                                        G_T = G_T + 1
                                    Case "V"    'Valine
                                        G_V = G_V + 1
                                    Case "W"    'Tryptophan
                                        G_W = G_W + 1
                                    Case "Y"    'Tyrosine
                                        G_Y = G_Y + 1
                                End Select
                            Case "I"    '5 helix (pi helix)
                                Select Case sequence(j).ToString
                                    Case "A"    'Alanine
                                        I_A = I_A + 1
                                    Case "C"    'Cysteine
                                        I_C = I_C + 1
                                    Case "D"    'Aspartate
                                        I_D = I_D + 1
                                    Case "E"    'Glutamate
                                        I_E = I_E + 1
                                    Case "F"    'Phenylalanine
                                        I_F = I_F + 1
                                    Case "G"    'Glycine
                                        I_G = I_G + 1
                                    Case "H"    'Histidine
                                        I_H = I_H + 1
                                    Case "I"    'Isoleucine
                                        I_I = I_I + 1
                                    Case "K"    'Lysine
                                        I_K = I_K + 1
                                    Case "L"    'Leucine
                                        I_L = I_L + 1
                                    Case "M"    'Methionine
                                        I_M = I_M + 1
                                    Case "N"    'Asparagine
                                        I_N = I_N + 1
                                    Case "P"    'Proline
                                        I_P = I_P + 1
                                    Case "Q"    'Glutamine
                                        I_Q = I_Q + 1
                                    Case "R"    'Arginine
                                        I_R = I_R + 1
                                    Case "S"    'Serine
                                        I_S = I_S + 1
                                    Case "T"    'Threonine
                                        I_T = I_T + 1
                                    Case "V"    'Valine
                                        I_V = I_V + 1
                                    Case "W"    'Tryptophan
                                        I_W = I_W + 1
                                    Case "Y"    'Tyrosine
                                        I_Y = I_Y + 1
                                End Select
                            Case "T"    'hydrogen bonded turn
                                Select Case sequence(j).ToString
                                    Case "A"    'Alanine
                                        T_A = T_A + 1
                                    Case "C"    'Cysteine
                                        T_C = T_C + 1
                                    Case "D"    'Aspartate
                                        T_D = T_D + 1
                                    Case "E"    'Glutamate
                                        T_E = T_E + 1
                                    Case "F"    'Phenylalanine
                                        T_F = T_F + 1
                                    Case "G"    'Glycine
                                        T_G = T_G + 1
                                    Case "H"    'Histidine
                                        T_H = T_H + 1
                                    Case "I"    'Isoleucine
                                        T_I = T_I + 1
                                    Case "K"    'Lysine
                                        T_K = T_K + 1
                                    Case "L"    'Leucine
                                        T_L = T_L + 1
                                    Case "M"    'Methionine
                                        T_M = T_M + 1
                                    Case "N"    'Asparagine
                                        T_N = T_N + 1
                                    Case "P"    'Proline
                                        T_P = T_P + 1
                                    Case "Q"    'Glutamine
                                        T_Q = T_Q + 1
                                    Case "R"    'Arginine
                                        T_R = T_R + 1
                                    Case "S"    'Serine
                                        T_S = T_S + 1
                                    Case "T"    'Threonine
                                        T_T = T_T + 1
                                    Case "V"    'Valine
                                        T_V = T_V + 1
                                    Case "W"    'Tryptophan
                                        T_W = T_W + 1
                                    Case "Y"    'Tyrosine
                                        T_Y = T_Y + 1
                                End Select
                            Case "S"    'bend
                                Select Case sequence(j).ToString
                                    Case "A"    'Alanine
                                        S_A = S_A + 1
                                    Case "C"    'Cysteine
                                        S_C = S_C + 1
                                    Case "D"    'Aspartate
                                        S_D = S_D + 1
                                    Case "E"    'Glutamate
                                        S_E = S_E + 1
                                    Case "F"    'Phenylalanine
                                        S_F = S_F + 1
                                    Case "G"    'Glycine
                                        S_G = S_G + 1
                                    Case "H"    'Histidine
                                        S_H = S_H + 1
                                    Case "I"    'Isoleucine
                                        S_I = S_I + 1
                                    Case "K"    'Lysine
                                        S_K = S_K + 1
                                    Case "L"    'Leucine
                                        S_L = S_L + 1
                                    Case "M"    'Methionine
                                        S_M = S_M + 1
                                    Case "N"    'Asparagine
                                        S_N = S_N + 1
                                    Case "P"    'Proline
                                        S_P = S_P + 1
                                    Case "Q"    'Glutamine
                                        S_Q = S_Q + 1
                                    Case "R"    'Arginine
                                        S_R = S_R + 1
                                    Case "S"    'Serine
                                        S_S = S_S + 1
                                    Case "T"    'Threonine
                                        S_T = S_T + 1
                                    Case "V"    'Valine
                                        S_V = S_V + 1
                                    Case "W"    'Tryptophan
                                        S_W = S_W + 1
                                    Case "Y"    'Tyrosine
                                        S_Y = S_Y + 1
                                End Select
                            Case " "    'loop or other irregular structure
                                'do nothing - just skip
                        End Select
                    Next
                End If
            Next
        End If

        Dim AlphaHelixCounter As Integer = H_A + H_C + H_D + H_E + H_F + H_G + H_H + H_I + H_K + H_L + H_M + H_N + H_P + H_Q + H_R + H_S + H_T + H_V + H_W + H_Y
        Dim BetaBridgeCounter As Integer = B_A + B_C + B_D + B_E + B_F + B_G + B_H + B_I + B_K + B_L + B_M + B_N + B_P + B_Q + B_R + B_S + B_T + B_V + B_W + B_Y
        Dim ExtendedStrandCounter As Integer = E_A + E_C + E_D + E_E + E_F + E_G + E_H + E_I + E_K + E_L + E_M + E_N + E_P + E_Q + E_R + E_S + E_T + E_V + E_W + E_Y
        Dim ThreeHelixCounter As Integer = G_A + G_C + G_D + G_E + G_F + G_G + G_H + G_I + G_K + G_L + G_M + G_N + G_P + G_Q + G_R + G_S + G_T + G_V + G_W + G_Y
        Dim FiveHelixCounter As Integer = I_A + I_C + I_D + I_E + I_F + I_G + I_H + I_I + I_K + I_L + I_M + I_N + I_P + I_Q + I_R + I_S + I_T + I_V + I_W + I_Y
        Dim HydroBondedTurnCounter As Integer = T_A + T_C + T_D + T_E + T_F + T_G + T_H + T_I + T_K + T_L + T_M + T_N + T_P + T_Q + T_R + T_S + T_T + T_V + T_W + T_Y
        Dim BendCounter As Integer = S_A + S_C + S_D + S_E + S_F + S_G + S_H + S_I + S_K + S_L + S_M + S_N + S_P + S_Q + S_R + S_S + S_T + S_V + S_W + S_Y

        Dim AlanineCounter As Integer = H_A + B_A + E_A + G_A + I_A + T_A + S_A
        Dim CysteineCounter As Integer = H_C + B_C + E_C + G_C + I_C + T_C + S_C
        Dim AspartateCounter As Integer = H_D + B_D + E_D + G_D + I_D + T_D + S_D
        Dim GlutamateCounter As Integer = H_E + B_E + E_E + G_E + I_E + T_E + S_E
        Dim PhenylalanineCounter As Integer = H_F + B_F + E_F + G_F + I_F + T_F + S_F
        Dim GlycineCounter As Integer = H_G + B_G + E_G + G_G + I_G + T_G + S_G
        Dim HistidineCounter As Integer = H_H + B_H + E_H + G_H + I_H + T_H + S_H
        Dim IsoleucineCounter As Integer = H_I + B_I + E_I + G_I + I_I + T_I + S_I
        Dim LysineCounter As Integer = H_K + B_K + E_K + G_K + I_K + T_K + S_K
        Dim LeucineCounter As Integer = H_L + B_L + E_L + G_L + I_L + T_L + S_L
        Dim MethionineCounter As Integer = H_M + B_M + E_M + G_M + I_M + T_M + S_M
        Dim AsparagineCounter As Integer = H_N + B_N + E_N + G_N + I_N + T_N + S_N
        Dim ProlineCounter As Integer = H_P + B_P + E_P + G_P + I_P + T_P + S_P
        Dim GlutamineCounter As Integer = H_Q + B_Q + E_Q + G_Q + I_Q + T_Q + S_Q
        Dim ArginineCounter As Integer = H_R + B_R + E_R + G_R + I_R + T_R + S_R
        Dim SerineCounter As Integer = H_S + B_S + E_S + G_S + I_S + T_S + S_S
        Dim ThreonineCounter As Integer = H_T + B_T + E_T + G_T + I_T + T_T + S_T
        Dim ValineCounter As Integer = H_V + B_V + E_V + G_V + I_V + T_V + S_V
        Dim TryptophanCounter As Integer = H_W + B_W + E_W + G_W + I_W + T_W + S_W
        Dim TyrosineCounter As Integer = H_Y + B_Y + E_Y + G_Y + I_Y + T_Y + S_Y


        ''Print Sequence Quantities
        Console.WriteLine(vbCrLf & vbCrLf & "--------- Sequence Results ---------")
        Console.WriteLine(vbCrLf & sumcounter & " sequence characters were processed." & vbCrLf)
        Console.WriteLine("1.  Alanine - Ala - A: " & Acounter & " characters - " & Math.Round(Acounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("2.  Cysteine - Cys - C: " & Ccounter & " characters - " & Math.Round(Ccounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("3.  Aspartate - Asp - D: " & Dcounter & " characters - " & Math.Round(Dcounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("4.  Glutamate - Glu - E: " & Ecounter & " characters - " & Math.Round(Ecounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("5.  Phenylalanine - Phe - F: " & Fcounter & " characters - " & Math.Round(Fcounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("6.  Glycine - Gly - G: " & Gcounter & " characters - " & Math.Round(Gcounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("7.  Histidine - His - H: " & Hcounter & " characters - " & Math.Round(Hcounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("8.  Isoleucine - Ile - I: " & Icounter & " characters - " & Math.Round(Icounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("9.  Lysine - Lys - K: " & Kcounter & " characters - " & Math.Round(Kcounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("10. Leucine - Leu - L: " & Lcounter & " characters - " & Math.Round(Lcounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("11. Methionine - Met - M: " & Mcounter & " characters - " & Math.Round(Mcounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("12. Asparagine - Asn - N: " & Ncounter & " characters - " & Math.Round(Ncounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("13. Proline - Pro - P: " & Pcounter & " characters - " & Math.Round(Pcounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("14. Glutamine - Gln - Q: " & Qcounter & " characters - " & Math.Round(Qcounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("15. Arginine - Arg - R: " & Rcounter & " characters - " & Math.Round(Rcounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("16. Serine - Ser - S: " & Scounter & " characters - " & Math.Round(Scounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("17. Threonine - Thr - T: " & Tcounter & " characters - " & Math.Round(Tcounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("18. Valine - Val - V: " & Vcounter & " characters - " & Math.Round(Vcounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("19. Tryptophan - Trp - W: " & Wcounter & " characters - " & Math.Round(Wcounter / sumcounter * 100, 2) & " %")
        Console.WriteLine("20. Tyrosine - Tyr - Y: " & Ycounter & " characters - " & Math.Round(Ycounter / sumcounter * 100, 2) & " %" & vbCrLf)

        If skippedcounter > 0 Then
            Console.WriteLine(vbCrLf & "The following " & skippedcounter & " sequence characters were skipped." & vbCrLf)
            Console.WriteLine("'B': " & Bcounter & " characters - Aspartic acid (D) or Asparagine (N)")
            Console.WriteLine("'J': " & Jcounter & " characters - Leucine (L) or Isoleucine (I)")
            Console.WriteLine("'O': " & Ocounter & " characters - Pyrrolysine")
            Console.WriteLine("'U': " & Ucounter & " characters - Selenocysteine")
            Console.WriteLine("'X': " & Xcounter & " characters - Any Amino Acid")
            Console.WriteLine("'Z': " & Zcounter & " characters - Glutamic acid (E) or Glutamine (Q)" & vbCrLf)
        End If


        ''Print Sequence - Secondary Structure Results
        'Console.WriteLine(vbCrLf & vbCrLf & "--------- Sequence - Secondary Structure Results --------- ")


        ''Export Results
        Try
            'create the 1st worksheet
            ws = pck.Workbook.Worksheets.Add("Amino Acid Quantities")
            ws.Cells.AutoFitColumns()
            ws.Cells.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            ws.Cells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White)

            ws.Cells("A1").Value = "Date: " & Now.ToShortDateString
            ws.Cells("A2").Value = "Report: Amino Acid Quantities"
            ws.Cells("A3").Value = "Processed Amino Acids: " + sumcounter.ToString
            ws.Cells("A1:A3").Style.Font.Bold = True

            'ws.Cells.AutoFitColumns(25)
            ws.Column(1).Width = 35
            ws.Column(2).Width = 20
            ws.Column(3).Width = 20

            'header
            ws.Cells(5, 1, 5, 3).Style.Font.Bold = True
            ws.Cells(5, 1, 5, 3).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
            ws.Cells(5, 1, 5, 3).Style.WrapText() = True
            ws.Cells(5, 1, 5, 3).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center
            ws.Cells(5, 1, 5, 3).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            ws.Cells(5, 1, 5, 3).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 192, 192, 192))

            'borders
            ws.Cells(5, 1, 25, 3).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
            ws.Cells(5, 1, 25, 3).Style.Border.Bottom.Style = Style.ExcelBorderStyle.Thin
            ws.Cells(5, 1, 25, 3).Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
            ws.Cells(5, 1, 25, 3).Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
            ws.Cells(5, 1, 25, 3).Style.Border.Right.Style = Style.ExcelBorderStyle.Thin
            ws.Cells(5, 1, 25, 3).Style.Border.BorderAround(Style.ExcelBorderStyle.Thick)

            ws.Cells("A5").Value = "Amino Acid"
            ws.Cells("B5").Value = "Total Number of A/A"
            ws.Cells("C5").Value = "Percentage (%)"


            ws.Cells("A6").Value = "Alanine - Ala - A"
            ws.Cells("B6").Value = Acounter.ToString
            ws.Cells("C6").Value = Math.Round(Acounter / sumcounter * 100, 2).ToString

            ws.Cells("A7").Value = "Cysteine - Cys - C"
            ws.Cells("B7").Value = Ccounter.ToString
            ws.Cells("C7").Value = Math.Round(Ccounter / sumcounter * 100, 2).ToString

            ws.Cells("A8").Value = "Aspartate - Asp - D"
            ws.Cells("B8").Value = Dcounter.ToString
            ws.Cells("C8").Value = Math.Round(Dcounter / sumcounter * 100, 2).ToString

            ws.Cells("A9").Value = "Glutamate - Glu - E"
            ws.Cells("B9").Value = Ecounter.ToString
            ws.Cells("C9").Value = Math.Round(Ecounter / sumcounter * 100, 2).ToString

            ws.Cells("A10").Value = "Phenylalanine - Phe - F"
            ws.Cells("B10").Value = Fcounter.ToString
            ws.Cells("C10").Value = Math.Round(Fcounter / sumcounter * 100, 2).ToString

            ws.Cells("A11").Value = "Glycine - Gly - G"
            ws.Cells("B11").Value = Gcounter.ToString
            ws.Cells("C11").Value = Math.Round(Gcounter / sumcounter * 100, 2).ToString

            ws.Cells("A12").Value = "Histidine - His - H"
            ws.Cells("B12").Value = Hcounter.ToString
            ws.Cells("C12").Value = Math.Round(Hcounter / sumcounter * 100, 2).ToString

            ws.Cells("A13").Value = "Isoleucine - Ile - I"
            ws.Cells("B13").Value = Icounter.ToString
            ws.Cells("C13").Value = Math.Round(Icounter / sumcounter * 100, 2).ToString

            ws.Cells("A14").Value = "Lysine - Lys - K"
            ws.Cells("B14").Value = Kcounter.ToString
            ws.Cells("C14").Value = Math.Round(Kcounter / sumcounter * 100, 2).ToString

            ws.Cells("A15").Value = "Leucine - Leu - L"
            ws.Cells("B15").Value = Lcounter.ToString
            ws.Cells("C15").Value = Math.Round(Lcounter / sumcounter * 100, 2).ToString

            ws.Cells("A16").Value = "Methionine - Met - M"
            ws.Cells("B16").Value = Mcounter.ToString
            ws.Cells("C16").Value = Math.Round(Mcounter / sumcounter * 100, 2).ToString

            ws.Cells("A17").Value = "Asparagine - Asn - N"
            ws.Cells("B17").Value = Ncounter.ToString
            ws.Cells("C17").Value = Math.Round(Ncounter / sumcounter * 100, 2).ToString

            ws.Cells("A18").Value = "Proline - Pro - P"
            ws.Cells("B18").Value = Pcounter.ToString
            ws.Cells("C18").Value = Math.Round(Pcounter / sumcounter * 100, 2).ToString

            ws.Cells("A19").Value = "Glutamine - Gln - Q"
            ws.Cells("B19").Value = Qcounter.ToString
            ws.Cells("C19").Value = Math.Round(Qcounter / sumcounter * 100, 2).ToString

            ws.Cells("A20").Value = "Arginine - Arg - R"
            ws.Cells("B20").Value = Rcounter.ToString
            ws.Cells("C20").Value = Math.Round(Rcounter / sumcounter * 100, 2).ToString

            ws.Cells("A21").Value = "Serine - Ser - S"
            ws.Cells("B21").Value = Scounter.ToString
            ws.Cells("C21").Value = Math.Round(Scounter / sumcounter * 100, 2).ToString

            ws.Cells("A22").Value = "Threonine - Thr - T"
            ws.Cells("B22").Value = Tcounter.ToString
            ws.Cells("C22").Value = Math.Round(Tcounter / sumcounter * 100, 2).ToString

            ws.Cells("A23").Value = "Valine - Val - V"
            ws.Cells("B23").Value = Vcounter.ToString
            ws.Cells("C23").Value = Math.Round(Vcounter / sumcounter * 100, 2).ToString

            ws.Cells("A24").Value = "Tryptophan - Trp - W"
            ws.Cells("B24").Value = Wcounter.ToString
            ws.Cells("C24").Value = Math.Round(Wcounter / sumcounter * 100, 2).ToString

            ws.Cells("A25").Value = "Tyrosine - Tyr - Y"
            ws.Cells("B25").Value = Ycounter.ToString
            ws.Cells("C25").Value = Math.Round(Ycounter / sumcounter * 100, 2).ToString

            If skippedcounter > 0 Then

                ws.Cells("A27").Value = "Skipped Amino Acids: " + skippedcounter.ToString
                ws.Cells("A27").Style.Font.Bold = True

                'header
                ws.Cells(29, 1, 29, 3).Style.Font.Bold = True
                ws.Cells(29, 1, 29, 3).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                ws.Cells(29, 1, 29, 3).Style.WrapText() = True
                ws.Cells(29, 1, 29, 3).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center
                ws.Cells(29, 1, 29, 3).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                ws.Cells(29, 1, 29, 3).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 192, 192, 192))

                'borders
                ws.Cells(29, 1, 35, 3).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                ws.Cells(29, 1, 35, 3).Style.Border.Bottom.Style = Style.ExcelBorderStyle.Thin
                ws.Cells(29, 1, 35, 3).Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
                ws.Cells(29, 1, 35, 3).Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
                ws.Cells(29, 1, 35, 3).Style.Border.Right.Style = Style.ExcelBorderStyle.Thin
                ws.Cells(29, 1, 35, 3).Style.Border.BorderAround(Style.ExcelBorderStyle.Thick)

                ws.Cells("A29").Value = "Ambiguous amino acids"
                ws.Cells("B29").Value = "Letter Code"
                ws.Cells("C29").Value = "Total Characters Found"

                ws.Cells("A30").Value = "Aspartic acid (D) or Asparagine (N)"
                ws.Cells("B30").Value = "B"
                ws.Cells("C30").Value = Bcounter.ToString

                ws.Cells("A31").Value = "Leucine (L) or Isoleucine (I)"
                ws.Cells("B31").Value = "J"
                ws.Cells("C31").Value = Jcounter.ToString

                ws.Cells("A32").Value = "Pyrrolysine"
                ws.Cells("B32").Value = "O"
                ws.Cells("C32").Value = Ocounter.ToString

                ws.Cells("A33").Value = "Selenocysteine"
                ws.Cells("B33").Value = "U"
                ws.Cells("C33").Value = Ucounter.ToString

                ws.Cells("A34").Value = "Any Amino Acid"
                ws.Cells("B34").Value = "X"
                ws.Cells("C34").Value = Xcounter.ToString

                ws.Cells("A35").Value = "Glutamic acid (E) or Glutamine (Q)"
                ws.Cells("B35").Value = "Z"
                ws.Cells("C35").Value = Zcounter.ToString

            End If

            ''create the 2nd worksheet
            ws2 = pck.Workbook.Worksheets.Add("Amino Acids Ratio")
            ws2.Cells.AutoFitColumns()
            ws2.Cells.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            ws2.Cells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White)

            ws2.Cells("A1").Value = "Date: " & Now.ToShortDateString
            ws2.Cells("A2").Value = "Report: Amino Acids Ratio"
            ws2.Cells("A1:A2").Style.Font.Bold = True

            'ws.Cells.AutoFitColumns(25)
            ws2.Column(1).Width = 25
            ws2.Column(2).Width = 25
            ws2.Column(3).Width = 25
            ws2.Column(4).Width = 25
            ws2.Column(5).Width = 25
            ws2.Column(6).Width = 25
            ws2.Column(7).Width = 25
            ws2.Column(8).Width = 25

            'horizontal header
            ws2.Cells(4, 1, 4, 8).Style.Font.Bold = True
            ws2.Cells(4, 1, 4, 8).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
            ws2.Cells(4, 1, 4, 8).Style.WrapText() = True
            ws2.Cells(4, 1, 4, 8).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center
            ws2.Cells(4, 1, 4, 8).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            ws2.Cells(4, 1, 4, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 192, 192, 192))

            'vertical header
            ws2.Cells(4, 1, 24, 1).Style.Font.Bold = True
            ws2.Cells(4, 1, 24, 1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
            ws2.Cells(4, 1, 24, 1).Style.WrapText() = True
            ws2.Cells(4, 1, 24, 1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center
            ws2.Cells(4, 1, 24, 1).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            ws2.Cells(4, 1, 24, 1).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 192, 192, 192))

            ws2.Cells("A4").Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 192, 192, 192))

            'borders
            ws2.Cells(4, 1, 24, 8).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
            ws2.Cells(4, 1, 24, 8).Style.Border.Bottom.Style = Style.ExcelBorderStyle.Thin
            ws2.Cells(4, 1, 24, 8).Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
            ws2.Cells(4, 1, 24, 8).Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
            ws2.Cells(4, 1, 24, 8).Style.Border.Right.Style = Style.ExcelBorderStyle.Thin
            ws2.Cells(4, 1, 24, 8).Style.Border.BorderAround(Style.ExcelBorderStyle.Thick)

            'column colors
            ws2.Cells(5, 2, 24, 2).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))
            ws2.Cells(5, 3, 24, 3).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 213, 224, 215))
            ws2.Cells(5, 4, 24, 4).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))
            ws2.Cells(5, 5, 24, 5).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 213, 224, 215))
            ws2.Cells(5, 6, 24, 6).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))
            ws2.Cells(5, 7, 24, 7).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 213, 224, 215))
            ws2.Cells(5, 8, 24, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))

            ws2.Cells("A5").Value = "Alanine - Ala - A"
            ws2.Cells("A6").Value = "Cysteine - Cys - C"
            ws2.Cells("A7").Value = "Aspartate - Asp - D"
            ws2.Cells("A8").Value = "Glutamate - Glu - E"
            ws2.Cells("A9").Value = "Phenylalanine - Phe - F"
            ws2.Cells("A10").Value = "Glycine - Gly - G"
            ws2.Cells("A11").Value = "Histidine - His - H"
            ws2.Cells("A12").Value = "Isoleucine - Ile - I"
            ws2.Cells("A13").Value = "Lysine - Lys - K"
            ws2.Cells("A14").Value = "Leucine - Leu - L"
            ws2.Cells("A15").Value = "Methionine - Met - M"
            ws2.Cells("A16").Value = "Asparagine - Asn - N"
            ws2.Cells("A17").Value = "Proline - Pro - P"
            ws2.Cells("A18").Value = "Glutamine - Gln - Q"
            ws2.Cells("A19").Value = "Arginine - Arg - R"
            ws2.Cells("A20").Value = "Serine - Ser - S"
            ws2.Cells("A21").Value = "Threonine - Thr - T"
            ws2.Cells("A22").Value = "Valine - Val - V"
            ws2.Cells("A23").Value = "Tryptophan - Trp - W"
            ws2.Cells("A24").Value = "Tyrosine - Tyr - Y"


            ws2.Cells("B4").Value = "H - Alpha Helix"
            ws2.Cells("C4").Value = "B - Beta-Bridge"
            ws2.Cells("D4").Value = "E - Extended Strand"
            ws2.Cells("E4").Value = "G - 3-Helix"
            ws2.Cells("F4").Value = "I - 5-Helix"
            ws2.Cells("G4").Value = "T - Hydrogen Bonded Turn"
            ws2.Cells("H4").Value = "S - Bend"


            ''Fill with values
            'Alanine
            ws2.Cells("B5").Value = H_A.ToString & " - " & Math.Round(H_A / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C5").Value = B_A.ToString & " - " & Math.Round(B_A / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D5").Value = E_A.ToString & " - " & Math.Round(E_A / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E5").Value = G_A.ToString & " - " & Math.Round(G_A / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F5").Value = I_A.ToString & " - " & Math.Round(I_A / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G5").Value = T_A.ToString & " - " & Math.Round(T_A / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H5").Value = S_A.ToString & " - " & Math.Round(S_A / BendCounter * 100, 2).ToString & " %"


            'Cysteine
            ws2.Cells("B6").Value = H_C.ToString & " - " & Math.Round(H_C / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C6").Value = B_C.ToString & " - " & Math.Round(B_C / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D6").Value = E_C.ToString & " - " & Math.Round(E_C / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E6").Value = G_C.ToString & " - " & Math.Round(G_C / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F6").Value = I_C.ToString & " - " & Math.Round(I_C / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G6").Value = T_C.ToString & " - " & Math.Round(T_C / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H6").Value = S_C.ToString & " - " & Math.Round(S_C / BendCounter * 100, 2).ToString & " %"

            'Aspartate
            ws2.Cells("B7").Value = H_D.ToString & " - " & Math.Round(H_D / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C7").Value = B_D.ToString & " - " & Math.Round(B_D / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D7").Value = E_D.ToString & " - " & Math.Round(E_D / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E7").Value = G_D.ToString & " - " & Math.Round(G_D / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F7").Value = I_D.ToString & " - " & Math.Round(I_D / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G7").Value = T_D.ToString & " - " & Math.Round(T_D / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H7").Value = S_D.ToString & " - " & Math.Round(S_D / BendCounter * 100, 2).ToString & " %"

            'Glutamate
            ws2.Cells("B8").Value = H_E.ToString & " - " & Math.Round(H_E / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C8").Value = B_E.ToString & " - " & Math.Round(B_E / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D8").Value = E_E.ToString & " - " & Math.Round(E_E / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E8").Value = G_E.ToString & " - " & Math.Round(G_E / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F8").Value = I_E.ToString & " - " & Math.Round(I_E / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G8").Value = T_E.ToString & " - " & Math.Round(T_E / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H8").Value = S_E.ToString & " - " & Math.Round(S_E / BendCounter * 100, 2).ToString & " %"

            'Phenylalanine
            ws2.Cells("B9").Value = H_F.ToString & " - " & Math.Round(H_F / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C9").Value = B_F.ToString & " - " & Math.Round(B_F / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D9").Value = E_F.ToString & " - " & Math.Round(E_F / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E9").Value = G_F.ToString & " - " & Math.Round(G_F / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F9").Value = I_F.ToString & " - " & Math.Round(I_F / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G9").Value = T_F.ToString & " - " & Math.Round(T_F / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H9").Value = S_F.ToString & " - " & Math.Round(S_F / BendCounter * 100, 2).ToString & " %"

            'Glycine
            ws2.Cells("B10").Value = H_G.ToString & " - " & Math.Round(H_G / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C10").Value = B_G.ToString & " - " & Math.Round(B_G / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D10").Value = E_G.ToString & " - " & Math.Round(E_G / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E10").Value = G_G.ToString & " - " & Math.Round(G_G / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F10").Value = I_G.ToString & " - " & Math.Round(I_G / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G10").Value = T_G.ToString & " - " & Math.Round(T_G / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H10").Value = S_G.ToString & " - " & Math.Round(S_G / BendCounter * 100, 2).ToString & " %"

            'Histidine
            ws2.Cells("B11").Value = H_H.ToString & " - " & Math.Round(H_H / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C11").Value = B_H.ToString & " - " & Math.Round(B_H / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D11").Value = E_H.ToString & " - " & Math.Round(E_H / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E11").Value = G_H.ToString & " - " & Math.Round(G_H / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F11").Value = I_H.ToString & " - " & Math.Round(I_H / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G11").Value = T_H.ToString & " - " & Math.Round(T_H / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H11").Value = S_H.ToString & " - " & Math.Round(S_H / BendCounter * 100, 2).ToString & " %"

            'Isoleucine
            ws2.Cells("B12").Value = H_I.ToString & " - " & Math.Round(H_I / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C12").Value = B_I.ToString & " - " & Math.Round(B_I / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D12").Value = E_I.ToString & " - " & Math.Round(E_I / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E12").Value = G_I.ToString & " - " & Math.Round(G_I / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F12").Value = I_I.ToString & " - " & Math.Round(I_I / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G12").Value = T_I.ToString & " - " & Math.Round(T_I / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H12").Value = S_I.ToString & " - " & Math.Round(S_I / BendCounter * 100, 2).ToString & " %"

            'Lysine
            ws2.Cells("B13").Value = H_K.ToString & " - " & Math.Round(H_K / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C13").Value = B_K.ToString & " - " & Math.Round(B_K / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D13").Value = E_K.ToString & " - " & Math.Round(E_K / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E13").Value = G_K.ToString & " - " & Math.Round(G_K / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F13").Value = I_K.ToString & " - " & Math.Round(I_K / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G13").Value = T_K.ToString & " - " & Math.Round(T_K / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H13").Value = S_K.ToString & " - " & Math.Round(S_K / BendCounter * 100, 2).ToString & " %"

            'Leucine
            ws2.Cells("B14").Value = H_L.ToString & " - " & Math.Round(H_L / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C14").Value = B_L.ToString & " - " & Math.Round(B_L / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D14").Value = E_L.ToString & " - " & Math.Round(E_L / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E14").Value = G_L.ToString & " - " & Math.Round(G_L / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F14").Value = I_L.ToString & " - " & Math.Round(I_L / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G14").Value = T_L.ToString & " - " & Math.Round(T_L / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H14").Value = S_L.ToString & " - " & Math.Round(S_L / BendCounter * 100, 2).ToString & " %"

            'Methionine
            ws2.Cells("B15").Value = H_M.ToString & " - " & Math.Round(H_M / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C15").Value = B_M.ToString & " - " & Math.Round(B_M / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D15").Value = E_M.ToString & " - " & Math.Round(E_M / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E15").Value = G_M.ToString & " - " & Math.Round(G_M / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F15").Value = I_M.ToString & " - " & Math.Round(I_M / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G15").Value = T_M.ToString & " - " & Math.Round(T_M / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H15").Value = S_M.ToString & " - " & Math.Round(S_M / BendCounter * 100, 2).ToString & " %"

            'Asparagine
            ws2.Cells("B16").Value = H_N.ToString & " - " & Math.Round(H_N / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C16").Value = B_N.ToString & " - " & Math.Round(B_N / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D16").Value = E_N.ToString & " - " & Math.Round(E_N / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E16").Value = G_N.ToString & " - " & Math.Round(G_N / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F16").Value = I_N.ToString & " - " & Math.Round(I_N / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G16").Value = T_N.ToString & " - " & Math.Round(T_N / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H16").Value = S_N.ToString & " - " & Math.Round(S_N / BendCounter * 100, 2).ToString & " %"

            'Proline
            ws2.Cells("B17").Value = H_P.ToString & " - " & Math.Round(H_P / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C17").Value = B_P.ToString & " - " & Math.Round(B_P / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D17").Value = E_P.ToString & " - " & Math.Round(E_P / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E17").Value = G_P.ToString & " - " & Math.Round(G_P / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F17").Value = I_P.ToString & " - " & Math.Round(I_P / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G17").Value = T_P.ToString & " - " & Math.Round(T_P / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H17").Value = S_P.ToString & " - " & Math.Round(S_P / BendCounter * 100, 2).ToString & " %"

            'Glutamine
            ws2.Cells("B18").Value = H_Q.ToString & " - " & Math.Round(H_Q / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C18").Value = B_Q.ToString & " - " & Math.Round(B_Q / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D18").Value = E_Q.ToString & " - " & Math.Round(E_Q / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E18").Value = G_Q.ToString & " - " & Math.Round(G_Q / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F18").Value = I_Q.ToString & " - " & Math.Round(I_Q / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G18").Value = T_Q.ToString & " - " & Math.Round(T_Q / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H18").Value = S_Q.ToString & " - " & Math.Round(S_Q / BendCounter * 100, 2).ToString & " %"

            'Arginine
            ws2.Cells("B19").Value = H_R.ToString & " - " & Math.Round(H_R / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C19").Value = B_R.ToString & " - " & Math.Round(B_R / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D19").Value = E_R.ToString & " - " & Math.Round(E_R / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E19").Value = G_R.ToString & " - " & Math.Round(G_R / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F19").Value = I_R.ToString & " - " & Math.Round(I_R / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G19").Value = T_R.ToString & " - " & Math.Round(T_R / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H19").Value = S_R.ToString & " - " & Math.Round(S_R / BendCounter * 100, 2).ToString & " %"

            'Serine
            ws2.Cells("B20").Value = H_S.ToString & " - " & Math.Round(H_S / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C20").Value = B_S.ToString & " - " & Math.Round(B_S / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D20").Value = E_S.ToString & " - " & Math.Round(E_S / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E20").Value = G_S.ToString & " - " & Math.Round(G_S / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F20").Value = I_S.ToString & " - " & Math.Round(I_S / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G20").Value = T_S.ToString & " - " & Math.Round(T_S / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H20").Value = S_S.ToString & " - " & Math.Round(S_S / BendCounter * 100, 2).ToString & " %"

            'Threonine
            ws2.Cells("B21").Value = H_T.ToString & " - " & Math.Round(H_T / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C21").Value = B_T.ToString & " - " & Math.Round(B_T / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D21").Value = E_T.ToString & " - " & Math.Round(E_T / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E21").Value = G_T.ToString & " - " & Math.Round(G_T / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F21").Value = I_T.ToString & " - " & Math.Round(I_T / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G21").Value = T_T.ToString & " - " & Math.Round(T_T / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H21").Value = S_T.ToString & " - " & Math.Round(S_T / BendCounter * 100, 2).ToString & " %"

            'Valine
            ws2.Cells("B22").Value = H_V.ToString & " - " & Math.Round(H_V / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C22").Value = B_V.ToString & " - " & Math.Round(B_V / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D22").Value = E_V.ToString & " - " & Math.Round(E_V / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E22").Value = G_V.ToString & " - " & Math.Round(G_V / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F22").Value = I_V.ToString & " - " & Math.Round(I_V / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G22").Value = T_V.ToString & " - " & Math.Round(T_V / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H22").Value = S_V.ToString & " - " & Math.Round(S_V / BendCounter * 100, 2).ToString & " %"

            'Tryptophan
            ws2.Cells("B23").Value = H_W.ToString & " - " & Math.Round(H_W / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C23").Value = B_W.ToString & " - " & Math.Round(B_W / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D23").Value = E_W.ToString & " - " & Math.Round(E_W / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E23").Value = G_W.ToString & " - " & Math.Round(G_W / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F23").Value = I_W.ToString & " - " & Math.Round(I_W / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G23").Value = T_W.ToString & " - " & Math.Round(T_W / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H23").Value = S_W.ToString & " - " & Math.Round(S_W / BendCounter * 100, 2).ToString & " %"

            'Tyrosine
            ws2.Cells("B24").Value = H_Y.ToString & " - " & Math.Round(H_Y / AlphaHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("C24").Value = B_Y.ToString & " - " & Math.Round(B_Y / BetaBridgeCounter * 100, 2).ToString & " %"
            ws2.Cells("D24").Value = E_Y.ToString & " - " & Math.Round(E_Y / ExtendedStrandCounter * 100, 2).ToString & " %"
            ws2.Cells("E24").Value = G_Y.ToString & " - " & Math.Round(G_Y / ThreeHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("F24").Value = I_Y.ToString & " - " & Math.Round(I_Y / FiveHelixCounter * 100, 2).ToString & " %"
            ws2.Cells("G24").Value = T_Y.ToString & " - " & Math.Round(T_Y / HydroBondedTurnCounter * 100, 2).ToString & " %"
            ws2.Cells("H24").Value = S_Y.ToString & " - " & Math.Round(S_Y / BendCounter * 100, 2).ToString & " %"

            'TOTALS OF ALPHA HELIX
            'ws2.Cells("B25").Value = AlphaHelixCounter.ToString & " - " & Math.Round((H_A + H_C + H_D + H_E + H_F + H_G + H_H + H_I + H_K + H_L + H_M + H_N + H_P + H_Q + H_R + H_S + H_T + H_V + H_W + H_Y) / AlphaHelixCounter * 100, 2).ToString & " %"


            ''create the 3rd worksheet
            ws3 = pck.Workbook.Worksheets.Add("Secondary Elements Ratio")
            ws3.Cells.AutoFitColumns()
            ws3.Cells.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            ws3.Cells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White)

            ws3.Cells("A1").Value = "Date: " & Now.ToShortDateString
            ws3.Cells("A2").Value = "Report: Secondary Elements Ratio"
            ws3.Cells("A1:A2").Style.Font.Bold = True

            'ws.Cells.AutoFitColumns(25)
            ws3.Column(1).Width = 25
            ws3.Column(2).Width = 25
            ws3.Column(3).Width = 25
            ws3.Column(4).Width = 25
            ws3.Column(5).Width = 25
            ws3.Column(6).Width = 25
            ws3.Column(7).Width = 25
            ws3.Column(8).Width = 25

            'horizontal header
            ws3.Cells(4, 1, 4, 8).Style.Font.Bold = True
            ws3.Cells(4, 1, 4, 8).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
            ws3.Cells(4, 1, 4, 8).Style.WrapText() = True
            ws3.Cells(4, 1, 4, 8).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center
            ws3.Cells(4, 1, 4, 8).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            ws3.Cells(4, 1, 4, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 192, 192, 192))

            'vertical header
            ws3.Cells(4, 1, 24, 1).Style.Font.Bold = True
            ws3.Cells(4, 1, 24, 1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
            ws3.Cells(4, 1, 24, 1).Style.WrapText() = True
            ws3.Cells(4, 1, 24, 1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center
            ws3.Cells(4, 1, 24, 1).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            ws3.Cells(4, 1, 24, 1).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 192, 192, 192))

            ws3.Cells("A4").Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 192, 192, 192))

            'borders
            ws3.Cells(4, 1, 24, 8).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
            ws3.Cells(4, 1, 24, 8).Style.Border.Bottom.Style = Style.ExcelBorderStyle.Thin
            ws3.Cells(4, 1, 24, 8).Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
            ws3.Cells(4, 1, 24, 8).Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
            ws3.Cells(4, 1, 24, 8).Style.Border.Right.Style = Style.ExcelBorderStyle.Thin
            ws3.Cells(4, 1, 24, 8).Style.Border.BorderAround(Style.ExcelBorderStyle.Thick)

            'row colors
            ws3.Cells(5, 2, 5, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))
            ws3.Cells(6, 2, 6, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 213, 224, 215))
            ws3.Cells(7, 2, 7, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))
            ws3.Cells(8, 2, 8, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 213, 224, 215))
            ws3.Cells(9, 2, 9, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))
            ws3.Cells(10, 2, 10, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 213, 224, 215))
            ws3.Cells(11, 2, 11, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))
            ws3.Cells(12, 2, 12, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 213, 224, 215))
            ws3.Cells(13, 2, 13, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))
            ws3.Cells(14, 2, 14, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 213, 224, 215))
            ws3.Cells(15, 2, 15, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))
            ws3.Cells(16, 2, 16, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 213, 224, 215))
            ws3.Cells(17, 2, 17, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))
            ws3.Cells(18, 2, 18, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 213, 224, 215))
            ws3.Cells(19, 2, 19, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))
            ws3.Cells(20, 2, 20, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 213, 224, 215))
            ws3.Cells(21, 2, 21, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))
            ws3.Cells(22, 2, 22, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 213, 224, 215))
            ws3.Cells(23, 2, 23, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 233, 244, 235))
            ws3.Cells(24, 2, 24, 8).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 213, 224, 215))


            ws3.Cells("A5").Value = "Alanine - Ala - A"
            ws3.Cells("A6").Value = "Cysteine - Cys - C"
            ws3.Cells("A7").Value = "Aspartate - Asp - D"
            ws3.Cells("A8").Value = "Glutamate - Glu - E"
            ws3.Cells("A9").Value = "Phenylalanine - Phe - F"
            ws3.Cells("A10").Value = "Glycine - Gly - G"
            ws3.Cells("A11").Value = "Histidine - His - H"
            ws3.Cells("A12").Value = "Isoleucine - Ile - I"
            ws3.Cells("A13").Value = "Lysine - Lys - K"
            ws3.Cells("A14").Value = "Leucine - Leu - L"
            ws3.Cells("A15").Value = "Methionine - Met - M"
            ws3.Cells("A16").Value = "Asparagine - Asn - N"
            ws3.Cells("A17").Value = "Proline - Pro - P"
            ws3.Cells("A18").Value = "Glutamine - Gln - Q"
            ws3.Cells("A19").Value = "Arginine - Arg - R"
            ws3.Cells("A20").Value = "Serine - Ser - S"
            ws3.Cells("A21").Value = "Threonine - Thr - T"
            ws3.Cells("A22").Value = "Valine - Val - V"
            ws3.Cells("A23").Value = "Tryptophan - Trp - W"
            ws3.Cells("A24").Value = "Tyrosine - Tyr - Y"


            ws3.Cells("B4").Value = "H - Alpha Helix"
            ws3.Cells("C4").Value = "B - Beta-Bridge"
            ws3.Cells("D4").Value = "E - Extended Strand"
            ws3.Cells("E4").Value = "G - 3-Helix"
            ws3.Cells("F4").Value = "I - 5-Helix"
            ws3.Cells("G4").Value = "T - Hydrogen Bonded Turn"
            ws3.Cells("H4").Value = "S - Bend"


            ''Fill with values
            'Alanine
            ws3.Cells("B5").Value = H_A.ToString & " - " & Math.Round(H_A / AlanineCounter * 100, 2).ToString & " %"
            ws3.Cells("C5").Value = B_A.ToString & " - " & Math.Round(B_A / AlanineCounter * 100, 2).ToString & " %"
            ws3.Cells("D5").Value = E_A.ToString & " - " & Math.Round(E_A / AlanineCounter * 100, 2).ToString & " %"
            ws3.Cells("E5").Value = G_A.ToString & " - " & Math.Round(G_A / AlanineCounter * 100, 2).ToString & " %"
            ws3.Cells("F5").Value = I_A.ToString & " - " & Math.Round(I_A / AlanineCounter * 100, 2).ToString & " %"
            ws3.Cells("G5").Value = T_A.ToString & " - " & Math.Round(T_A / AlanineCounter * 100, 2).ToString & " %"
            ws3.Cells("H5").Value = S_A.ToString & " - " & Math.Round(S_A / AlanineCounter * 100, 2).ToString & " %"

            'Cysteine
            ws3.Cells("B6").Value = H_C.ToString & " - " & Math.Round(H_C / CysteineCounter * 100, 2).ToString & " %"
            ws3.Cells("C6").Value = B_C.ToString & " - " & Math.Round(B_C / CysteineCounter * 100, 2).ToString & " %"
            ws3.Cells("D6").Value = E_C.ToString & " - " & Math.Round(E_C / CysteineCounter * 100, 2).ToString & " %"
            ws3.Cells("E6").Value = G_C.ToString & " - " & Math.Round(G_C / CysteineCounter * 100, 2).ToString & " %"
            ws3.Cells("F6").Value = I_C.ToString & " - " & Math.Round(I_C / CysteineCounter * 100, 2).ToString & " %"
            ws3.Cells("G6").Value = T_C.ToString & " - " & Math.Round(T_C / CysteineCounter * 100, 2).ToString & " %"
            ws3.Cells("H6").Value = S_C.ToString & " - " & Math.Round(S_C / CysteineCounter * 100, 2).ToString & " %"

            'Aspartate
            ws3.Cells("B7").Value = H_D.ToString & " - " & Math.Round(H_D / AspartateCounter * 100, 2).ToString & " %"
            ws3.Cells("C7").Value = B_D.ToString & " - " & Math.Round(B_D / AspartateCounter * 100, 2).ToString & " %"
            ws3.Cells("D7").Value = E_D.ToString & " - " & Math.Round(E_D / AspartateCounter * 100, 2).ToString & " %"
            ws3.Cells("E7").Value = G_D.ToString & " - " & Math.Round(G_D / AspartateCounter * 100, 2).ToString & " %"
            ws3.Cells("F7").Value = I_D.ToString & " - " & Math.Round(I_D / AspartateCounter * 100, 2).ToString & " %"
            ws3.Cells("G7").Value = T_D.ToString & " - " & Math.Round(T_D / AspartateCounter * 100, 2).ToString & " %"
            ws3.Cells("H7").Value = S_D.ToString & " - " & Math.Round(S_D / AspartateCounter * 100, 2).ToString & " %"

            'Glutamate
            ws3.Cells("B8").Value = H_E.ToString & " - " & Math.Round(H_E / GlutamateCounter * 100, 2).ToString & " %"
            ws3.Cells("C8").Value = B_E.ToString & " - " & Math.Round(B_E / GlutamateCounter * 100, 2).ToString & " %"
            ws3.Cells("D8").Value = E_E.ToString & " - " & Math.Round(E_E / GlutamateCounter * 100, 2).ToString & " %"
            ws3.Cells("E8").Value = G_E.ToString & " - " & Math.Round(G_E / GlutamateCounter * 100, 2).ToString & " %"
            ws3.Cells("F8").Value = I_E.ToString & " - " & Math.Round(I_E / GlutamateCounter * 100, 2).ToString & " %"
            ws3.Cells("G8").Value = T_E.ToString & " - " & Math.Round(T_E / GlutamateCounter * 100, 2).ToString & " %"
            ws3.Cells("H8").Value = S_E.ToString & " - " & Math.Round(S_E / GlutamateCounter * 100, 2).ToString & " %"

            'Phenylalanine
            ws3.Cells("B9").Value = H_F.ToString & " - " & Math.Round(H_F / PhenylalanineCounter * 100, 2).ToString & " %"
            ws3.Cells("C9").Value = B_F.ToString & " - " & Math.Round(B_F / PhenylalanineCounter * 100, 2).ToString & " %"
            ws3.Cells("D9").Value = E_F.ToString & " - " & Math.Round(E_F / PhenylalanineCounter * 100, 2).ToString & " %"
            ws3.Cells("E9").Value = G_F.ToString & " - " & Math.Round(G_F / PhenylalanineCounter * 100, 2).ToString & " %"
            ws3.Cells("F9").Value = I_F.ToString & " - " & Math.Round(I_F / PhenylalanineCounter * 100, 2).ToString & " %"
            ws3.Cells("G9").Value = T_F.ToString & " - " & Math.Round(T_F / PhenylalanineCounter * 100, 2).ToString & " %"
            ws3.Cells("H9").Value = S_F.ToString & " - " & Math.Round(S_F / PhenylalanineCounter * 100, 2).ToString & " %"

            'Glycine
            ws3.Cells("B10").Value = H_G.ToString & " - " & Math.Round(H_G / GlycineCounter * 100, 2).ToString & " %"
            ws3.Cells("C10").Value = B_G.ToString & " - " & Math.Round(B_G / GlycineCounter * 100, 2).ToString & " %"
            ws3.Cells("D10").Value = E_G.ToString & " - " & Math.Round(E_G / GlycineCounter * 100, 2).ToString & " %"
            ws3.Cells("E10").Value = G_G.ToString & " - " & Math.Round(G_G / GlycineCounter * 100, 2).ToString & " %"
            ws3.Cells("F10").Value = I_G.ToString & " - " & Math.Round(I_G / GlycineCounter * 100, 2).ToString & " %"
            ws3.Cells("G10").Value = T_G.ToString & " - " & Math.Round(T_G / GlycineCounter * 100, 2).ToString & " %"
            ws3.Cells("H10").Value = S_G.ToString & " - " & Math.Round(S_G / GlycineCounter * 100, 2).ToString & " %"

            'Histidine
            ws3.Cells("B11").Value = H_H.ToString & " - " & Math.Round(H_H / HistidineCounter * 100, 2).ToString & " %"
            ws3.Cells("C11").Value = B_H.ToString & " - " & Math.Round(B_H / HistidineCounter * 100, 2).ToString & " %"
            ws3.Cells("D11").Value = E_H.ToString & " - " & Math.Round(E_H / HistidineCounter * 100, 2).ToString & " %"
            ws3.Cells("E11").Value = G_H.ToString & " - " & Math.Round(G_H / HistidineCounter * 100, 2).ToString & " %"
            ws3.Cells("F11").Value = I_H.ToString & " - " & Math.Round(I_H / HistidineCounter * 100, 2).ToString & " %"
            ws3.Cells("G11").Value = T_H.ToString & " - " & Math.Round(T_H / HistidineCounter * 100, 2).ToString & " %"
            ws3.Cells("H11").Value = S_H.ToString & " - " & Math.Round(S_H / HistidineCounter * 100, 2).ToString & " %"

            'Isoleucine
            ws3.Cells("B12").Value = H_I.ToString & " - " & Math.Round(H_I / IsoleucineCounter * 100, 2).ToString & " %"
            ws3.Cells("C12").Value = B_I.ToString & " - " & Math.Round(B_I / IsoleucineCounter * 100, 2).ToString & " %"
            ws3.Cells("D12").Value = E_I.ToString & " - " & Math.Round(E_I / IsoleucineCounter * 100, 2).ToString & " %"
            ws3.Cells("E12").Value = G_I.ToString & " - " & Math.Round(G_I / IsoleucineCounter * 100, 2).ToString & " %"
            ws3.Cells("F12").Value = I_I.ToString & " - " & Math.Round(I_I / IsoleucineCounter * 100, 2).ToString & " %"
            ws3.Cells("G12").Value = T_I.ToString & " - " & Math.Round(T_I / IsoleucineCounter * 100, 2).ToString & " %"
            ws3.Cells("H12").Value = S_I.ToString & " - " & Math.Round(S_I / IsoleucineCounter * 100, 2).ToString & " %"

            'Lysine
            ws3.Cells("B13").Value = H_K.ToString & " - " & Math.Round(H_K / LysineCounter * 100, 2).ToString & " %"
            ws3.Cells("C13").Value = B_K.ToString & " - " & Math.Round(B_K / LysineCounter * 100, 2).ToString & " %"
            ws3.Cells("D13").Value = E_K.ToString & " - " & Math.Round(E_K / LysineCounter * 100, 2).ToString & " %"
            ws3.Cells("E13").Value = G_K.ToString & " - " & Math.Round(G_K / LysineCounter * 100, 2).ToString & " %"
            ws3.Cells("F13").Value = I_K.ToString & " - " & Math.Round(I_K / LysineCounter * 100, 2).ToString & " %"
            ws3.Cells("G13").Value = T_K.ToString & " - " & Math.Round(T_K / LysineCounter * 100, 2).ToString & " %"
            ws3.Cells("H13").Value = S_K.ToString & " - " & Math.Round(S_K / LysineCounter * 100, 2).ToString & " %"

            'Leucine
            ws3.Cells("B14").Value = H_L.ToString & " - " & Math.Round(H_L / LeucineCounter * 100, 2).ToString & " %"
            ws3.Cells("C14").Value = B_L.ToString & " - " & Math.Round(B_L / LeucineCounter * 100, 2).ToString & " %"
            ws3.Cells("D14").Value = E_L.ToString & " - " & Math.Round(E_L / LeucineCounter * 100, 2).ToString & " %"
            ws3.Cells("E14").Value = G_L.ToString & " - " & Math.Round(G_L / LeucineCounter * 100, 2).ToString & " %"
            ws3.Cells("F14").Value = I_L.ToString & " - " & Math.Round(I_L / LeucineCounter * 100, 2).ToString & " %"
            ws3.Cells("G14").Value = T_L.ToString & " - " & Math.Round(T_L / LeucineCounter * 100, 2).ToString & " %"
            ws3.Cells("H14").Value = S_L.ToString & " - " & Math.Round(S_L / LeucineCounter * 100, 2).ToString & " %"

            'Methionine
            ws3.Cells("B15").Value = H_M.ToString & " - " & Math.Round(H_M / MethionineCounter * 100, 2).ToString & " %"
            ws3.Cells("C15").Value = B_M.ToString & " - " & Math.Round(B_M / MethionineCounter * 100, 2).ToString & " %"
            ws3.Cells("D15").Value = E_M.ToString & " - " & Math.Round(E_M / MethionineCounter * 100, 2).ToString & " %"
            ws3.Cells("E15").Value = G_M.ToString & " - " & Math.Round(G_M / MethionineCounter * 100, 2).ToString & " %"
            ws3.Cells("F15").Value = I_M.ToString & " - " & Math.Round(I_M / MethionineCounter * 100, 2).ToString & " %"
            ws3.Cells("G15").Value = T_M.ToString & " - " & Math.Round(T_M / MethionineCounter * 100, 2).ToString & " %"
            ws3.Cells("H15").Value = S_M.ToString & " - " & Math.Round(S_M / MethionineCounter * 100, 2).ToString & " %"

            'Asparagine
            ws3.Cells("B16").Value = H_N.ToString & " - " & Math.Round(H_N / AsparagineCounter * 100, 2).ToString & " %"
            ws3.Cells("C16").Value = B_N.ToString & " - " & Math.Round(B_N / AsparagineCounter * 100, 2).ToString & " %"
            ws3.Cells("D16").Value = E_N.ToString & " - " & Math.Round(E_N / AsparagineCounter * 100, 2).ToString & " %"
            ws3.Cells("E16").Value = G_N.ToString & " - " & Math.Round(G_N / AsparagineCounter * 100, 2).ToString & " %"
            ws3.Cells("F16").Value = I_N.ToString & " - " & Math.Round(I_N / AsparagineCounter * 100, 2).ToString & " %"
            ws3.Cells("G16").Value = T_N.ToString & " - " & Math.Round(T_N / AsparagineCounter * 100, 2).ToString & " %"
            ws3.Cells("H16").Value = S_N.ToString & " - " & Math.Round(S_N / AsparagineCounter * 100, 2).ToString & " %"

            'Proline
            ws3.Cells("B17").Value = H_P.ToString & " - " & Math.Round(H_P / ProlineCounter * 100, 2).ToString & " %"
            ws3.Cells("C17").Value = B_P.ToString & " - " & Math.Round(B_P / ProlineCounter * 100, 2).ToString & " %"
            ws3.Cells("D17").Value = E_P.ToString & " - " & Math.Round(E_P / ProlineCounter * 100, 2).ToString & " %"
            ws3.Cells("E17").Value = G_P.ToString & " - " & Math.Round(G_P / ProlineCounter * 100, 2).ToString & " %"
            ws3.Cells("F17").Value = I_P.ToString & " - " & Math.Round(I_P / ProlineCounter * 100, 2).ToString & " %"
            ws3.Cells("G17").Value = T_P.ToString & " - " & Math.Round(T_P / ProlineCounter * 100, 2).ToString & " %"
            ws3.Cells("H17").Value = S_P.ToString & " - " & Math.Round(S_P / ProlineCounter * 100, 2).ToString & " %"

            'Glutamine
            ws3.Cells("B18").Value = H_Q.ToString & " - " & Math.Round(H_Q / GlutamineCounter * 100, 2).ToString & " %"
            ws3.Cells("C18").Value = B_Q.ToString & " - " & Math.Round(B_Q / GlutamineCounter * 100, 2).ToString & " %"
            ws3.Cells("D18").Value = E_Q.ToString & " - " & Math.Round(E_Q / GlutamineCounter * 100, 2).ToString & " %"
            ws3.Cells("E18").Value = G_Q.ToString & " - " & Math.Round(G_Q / GlutamineCounter * 100, 2).ToString & " %"
            ws3.Cells("F18").Value = I_Q.ToString & " - " & Math.Round(I_Q / GlutamineCounter * 100, 2).ToString & " %"
            ws3.Cells("G18").Value = T_Q.ToString & " - " & Math.Round(T_Q / GlutamineCounter * 100, 2).ToString & " %"
            ws3.Cells("H18").Value = S_Q.ToString & " - " & Math.Round(S_Q / GlutamineCounter * 100, 2).ToString & " %"

            'Arginine
            ws3.Cells("B19").Value = H_R.ToString & " - " & Math.Round(H_R / ArginineCounter * 100, 2).ToString & " %"
            ws3.Cells("C19").Value = B_R.ToString & " - " & Math.Round(B_R / ArginineCounter * 100, 2).ToString & " %"
            ws3.Cells("D19").Value = E_R.ToString & " - " & Math.Round(E_R / ArginineCounter * 100, 2).ToString & " %"
            ws3.Cells("E19").Value = G_R.ToString & " - " & Math.Round(G_R / ArginineCounter * 100, 2).ToString & " %"
            ws3.Cells("F19").Value = I_R.ToString & " - " & Math.Round(I_R / ArginineCounter * 100, 2).ToString & " %"
            ws3.Cells("G19").Value = T_R.ToString & " - " & Math.Round(T_R / ArginineCounter * 100, 2).ToString & " %"
            ws3.Cells("H19").Value = S_R.ToString & " - " & Math.Round(S_R / ArginineCounter * 100, 2).ToString & " %"

            'Serine
            ws3.Cells("B20").Value = H_S.ToString & " - " & Math.Round(H_S / SerineCounter * 100, 2).ToString & " %"
            ws3.Cells("C20").Value = B_S.ToString & " - " & Math.Round(B_S / SerineCounter * 100, 2).ToString & " %"
            ws3.Cells("D20").Value = E_S.ToString & " - " & Math.Round(E_S / SerineCounter * 100, 2).ToString & " %"
            ws3.Cells("E20").Value = G_S.ToString & " - " & Math.Round(G_S / SerineCounter * 100, 2).ToString & " %"
            ws3.Cells("F20").Value = I_S.ToString & " - " & Math.Round(I_S / SerineCounter * 100, 2).ToString & " %"
            ws3.Cells("G20").Value = T_S.ToString & " - " & Math.Round(T_S / SerineCounter * 100, 2).ToString & " %"
            ws3.Cells("H20").Value = S_S.ToString & " - " & Math.Round(S_S / SerineCounter * 100, 2).ToString & " %"

            'Threonine
            ws3.Cells("B21").Value = H_T.ToString & " - " & Math.Round(H_T / ThreonineCounter * 100, 2).ToString & " %"
            ws3.Cells("C21").Value = B_T.ToString & " - " & Math.Round(B_T / ThreonineCounter * 100, 2).ToString & " %"
            ws3.Cells("D21").Value = E_T.ToString & " - " & Math.Round(E_T / ThreonineCounter * 100, 2).ToString & " %"
            ws3.Cells("E21").Value = G_T.ToString & " - " & Math.Round(G_T / ThreonineCounter * 100, 2).ToString & " %"
            ws3.Cells("F21").Value = I_T.ToString & " - " & Math.Round(I_T / ThreonineCounter * 100, 2).ToString & " %"
            ws3.Cells("G21").Value = T_T.ToString & " - " & Math.Round(T_T / ThreonineCounter * 100, 2).ToString & " %"
            ws3.Cells("H21").Value = S_T.ToString & " - " & Math.Round(S_T / ThreonineCounter * 100, 2).ToString & " %"

            'Valine
            ws3.Cells("B22").Value = H_V.ToString & " - " & Math.Round(H_V / ValineCounter * 100, 2).ToString & " %"
            ws3.Cells("C22").Value = B_V.ToString & " - " & Math.Round(B_V / ValineCounter * 100, 2).ToString & " %"
            ws3.Cells("D22").Value = E_V.ToString & " - " & Math.Round(E_V / ValineCounter * 100, 2).ToString & " %"
            ws3.Cells("E22").Value = G_V.ToString & " - " & Math.Round(G_V / ValineCounter * 100, 2).ToString & " %"
            ws3.Cells("F22").Value = I_V.ToString & " - " & Math.Round(I_V / ValineCounter * 100, 2).ToString & " %"
            ws3.Cells("G22").Value = T_V.ToString & " - " & Math.Round(T_V / ValineCounter * 100, 2).ToString & " %"
            ws3.Cells("H22").Value = S_V.ToString & " - " & Math.Round(S_V / ValineCounter * 100, 2).ToString & " %"

            'Tryptophan
            ws3.Cells("B23").Value = H_W.ToString & " - " & Math.Round(H_W / TryptophanCounter * 100, 2).ToString & " %"
            ws3.Cells("C23").Value = B_W.ToString & " - " & Math.Round(B_W / TryptophanCounter * 100, 2).ToString & " %"
            ws3.Cells("D23").Value = E_W.ToString & " - " & Math.Round(E_W / TryptophanCounter * 100, 2).ToString & " %"
            ws3.Cells("E23").Value = G_W.ToString & " - " & Math.Round(G_W / TryptophanCounter * 100, 2).ToString & " %"
            ws3.Cells("F23").Value = I_W.ToString & " - " & Math.Round(I_W / TryptophanCounter * 100, 2).ToString & " %"
            ws3.Cells("G23").Value = T_W.ToString & " - " & Math.Round(T_W / TryptophanCounter * 100, 2).ToString & " %"
            ws3.Cells("H23").Value = S_W.ToString & " - " & Math.Round(S_W / TryptophanCounter * 100, 2).ToString & " %"

            'Tyrosine
            ws3.Cells("B24").Value = H_Y.ToString & " - " & Math.Round(H_Y / TyrosineCounter * 100, 2).ToString & " %"
            ws3.Cells("C24").Value = B_Y.ToString & " - " & Math.Round(B_Y / TyrosineCounter * 100, 2).ToString & " %"
            ws3.Cells("D24").Value = E_Y.ToString & " - " & Math.Round(E_Y / TyrosineCounter * 100, 2).ToString & " %"
            ws3.Cells("E24").Value = G_Y.ToString & " - " & Math.Round(G_Y / TyrosineCounter * 100, 2).ToString & " %"
            ws3.Cells("F24").Value = I_Y.ToString & " - " & Math.Round(I_Y / TyrosineCounter * 100, 2).ToString & " %"
            ws3.Cells("G24").Value = T_Y.ToString & " - " & Math.Round(T_Y / TyrosineCounter * 100, 2).ToString & " %"
            ws3.Cells("H24").Value = S_Y.ToString & " - " & Math.Round(S_Y / TyrosineCounter * 100, 2).ToString & " %"


            pck.SaveAs(New FileInfo(path & "\SeqSec_Report.xlsx"))
            Console.WriteLine(vbCrLf & "A report file was generated successfully under the following path. " & vbCrLf & path)

        Catch ex As Exception
            Console.WriteLine("An error occured while exporting results.")
            Console.WriteLine("The error message is: " & ex.Message)
        End Try

    End Sub

End Module
