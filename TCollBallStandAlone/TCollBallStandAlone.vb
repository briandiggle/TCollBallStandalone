Imports System.Text
Imports System.IO
Imports System.Configuration

Module TCollBallStandAlone

    '======================CONVERT THIS APP TO USE THE COMMAND PATTERN=======================
    '----Populated from app.config------
    Dim CONSTANTS_FILE As String
    Dim FULL_DAY_UPDATE As Double
    Dim LUNCH_ONLY As Double
    Dim JOURNEYS_ONLY As Double
    Dim JOURNEY As Double
    Dim BUSES_LAST_INCREASE As String
    Dim LUNCH_LAST_INCREASE As String
    Dim BALANCE_FILE As String = "C:\Dev\TCollBall\BalFiles\tcball.txt"
    Dim BALANCE_FILE_BACKUP As String = "C:\Dev\TCollBall\BalFiles\tcball.bak"
    Dim SHOPPING_LIST_FILE As String

    Dim dBalance As Double = 0D
    Dim dToMove As Double = 0D
    Dim strLastUpdated As String = ""
    Dim strBalanceNote As String = ""

    Sub Main()

        If Not GetConstantsFromFile() Then
            Exit Sub
        End If

        Dim bFinished As Boolean = False
        Dim cOption As Char

        Dim strMessage As String = ""

        GetBalFile()


        DisplayWelcomeMessage()
        Console.WriteLine(ShowCurrentPosition())

        While Not bFinished
            cOption = ShowOptionsAndGetResponse()

            If cOption = "x" Then
                bFinished = True
            ElseIf cOption = "1" Then
                AddToBalance(FULL_DAY_UPDATE)
                Console.Write(ShowCurrentPosition)
            ElseIf cOption = "2" Then
                AddToBalance(LUNCH_ONLY)
                Console.Write(ShowCurrentPosition)
            ElseIf cOption = "3" Then
                AddToBalance(JOURNEYS_ONLY)
                Console.Write(ShowCurrentPosition)
            ElseIf cOption = "4" Then
                AddToBalance(JOURNEY)
                Console.Write(ShowCurrentPosition)
            ElseIf cOption = "5" Then
                Dim dValMoved As Double
                dValMoved = GetValueMoved()
                MovedFromTCAccount(dValMoved)
                Console.Write(ShowCurrentPosition)
            ElseIf cOption = "6" Then
                Dim dValSpent As Double
                dValSpent = GetValueSpent()
                Spent(dValSpent)
                Console.Write(ShowCurrentPosition)
            ElseIf cOption = "7" Then
                Dim strNewBalanceNote As String
                strNewBalanceNote = GetNewBalanceNote()
                ChangeBalanceNote(strNewBalanceNote)
                Console.Write(ShowCurrentPosition)
            ElseIf cOption = "8" Then
                Console.Write(ShowValues)
            ElseIf cOption = "9" Then
                Console.WriteLine(ShowCurrentPosition())
            ElseIf cOption = "0" Then
                Console.WriteLine(GetShoppingList())
            End If

        End While



    End Sub


    Public Function ShowOptionsAndGetResponse() As Char
        Dim bValidResponse As Boolean = False
        Dim oResponse As System.ConsoleKeyInfo
        Dim cResponse As Char = ""

        Console.WriteLine("Options")
        Console.WriteLine("1: Full day")
        Console.WriteLine("2: Lunch only")
        Console.WriteLine("3: Journeys only")
        Console.WriteLine("4: Single Journey only")
        Console.WriteLine("5: Update ToMove")
        Console.WriteLine("6: Update balance")
        Console.WriteLine("7: Balance Note")
        Console.WriteLine("8: Display Values")
        Console.WriteLine("9: Show Current Position")
        Console.WriteLine("0: Shopping List")
        Console.WriteLine("x: Finish")

        Console.WriteLine()

        While Not bValidResponse
            Console.Write("> ")
            oResponse = Console.ReadKey()
            Console.WriteLine()

            If oResponse.KeyChar.Equals("1"c) OrElse oResponse.KeyChar.Equals("0"c) OrElse oResponse.KeyChar.Equals("2"c) OrElse oResponse.KeyChar.Equals("3"c) OrElse oResponse.KeyChar.Equals("4"c) OrElse oResponse.KeyChar.Equals("5"c) OrElse oResponse.KeyChar.Equals("6"c) OrElse oResponse.KeyChar.Equals("7"c) OrElse oResponse.KeyChar.Equals("8"c) OrElse oResponse.KeyChar.Equals("9"c) OrElse oResponse.KeyChar.Equals("x"c) Then
                bValidResponse = True
                cResponse = oResponse.KeyChar
            End If

        End While

        Return cResponse

    End Function

    Public Sub DisplayWelcomeMessage()
        Dim oSB As StringBuilder = New StringBuilder

 
        oSB.AppendLine("TColl Balance Version: Standalone")
        oSB.AppendLine("=================================")


        Console.WriteLine(oSB.ToString)

    End Sub

    Public Function ShowValues() As String

        Dim oSB As StringBuilder = New StringBuilder


        oSB.AppendLine()

        oSB.AppendLine("Show Values:")
        oSB.AppendLine("------------")

        Try
            oSB.AppendLine("Full Day      : " + FULL_DAY_UPDATE.ToString("00.00"))
        Catch ex As Exception
            oSB.AppendLine("Full Day      : error here")
        End Try

        Try
            oSB.AppendLine("Lunch Only    : " + LUNCH_ONLY.ToString("00.00"))
        Catch ex As Exception
            oSB.AppendLine("Lunch Only    : error here")
        End Try

        Try
            oSB.AppendLine("Journeys Only : " + JOURNEYS_ONLY.ToString("00.00"))
        Catch ex As Exception
            oSB.AppendLine("Journeys Only : error here")
        End Try

        Try
            oSB.AppendLine("Journey       : " + JOURNEY.ToString("00.00"))
        Catch ex As Exception
            oSB.AppendLine("Journey       : error here")
        End Try

        oSB.AppendLine("Buses Last Increase : " + BUSES_LAST_INCREASE)
        oSB.AppendLine("Lunch Last Increase : " + LUNCH_LAST_INCREASE)

        oSB.AppendLine()

        oSB.AppendLine("Constants file at " + CONSTANTS_FILE)

        oSB.AppendLine()

        Return oSB.ToString

    End Function

    Public Function ShowCurrentPosition() As String

        Dim oSB As StringBuilder = New StringBuilder

        oSB.AppendLine()

        oSB.AppendLine("Current Position:")
        oSB.AppendLine("----------------")

        Try

            oSB.AppendLine("Balance      : " + dBalance.ToString("00.00"))
        Catch ex As Exception
            oSB.AppendLine("Balance      : error here")
        End Try

        oSB.AppendLine("Balance note : " + strBalanceNote)

        Try
            oSB.AppendLine("To Move      : " + dToMove.ToString())
        Catch ex As Exception
            oSB.AppendLine("To Move      : error here")
        End Try

        oSB.AppendLine("Last Update  : " + strLastUpdated)
        oSB.AppendLine()
        Return oSB.ToString

    End Function

    Public Sub GetBalFile()

        Using oReader As New StreamReader(BALANCE_FILE)
            Dim strLine As String

            '----First line holds the balance-------
            strLine = oReader.ReadLine
            dBalance = Double.Parse(strLine)

            '----Second Line holds the balance note-----
            strBalanceNote = oReader.ReadLine

            '----Third line holds to move----------
            strLine = oReader.ReadLine
            dToMove = Double.Parse(strLine)

            '----Fourth line holds the last updated string-----
            strLastUpdated = oReader.ReadLine

            oReader.Close()
        End Using


    End Sub


    Public Function AddToBalance(ByVal dValToAdd As Double) As String
        dBalance = dBalance + dValToAdd
        dToMove = dToMove + dValToAdd
        SaveBalance()

        Return "Updated balance is " + dBalance.ToString
    End Function

    Public Function MovedFromTCAccount(ByVal dValMoved As Double)
        dToMove = dToMove - dValMoved
        SaveBalance()
        Return "Updated to move is " + dToMove.ToString
    End Function

    Public Function Spent(ByVal dValSpent As Double) As String
        dBalance = dBalance - dValSpent
        SaveBalance()

        Return "Updated balance is " + dBalance.ToString
    End Function

    Public Function ChangeBalanceNote(ByVal strNewBalanceNote As String) As String

        strBalanceNote = strNewBalanceNote
        SaveBalance()
        Return "Updated to balance note " + dToMove.ToString
    End Function

    Private Function SaveBalance() As Integer

        Dim oWriter As StreamWriter

        '----Save a copy of the existing balance file-----------------------
    
        Try
            If File.Exists(BALANCE_FILE_BACKUP) Then
                File.Move(BALANCE_FILE_BACKUP, BALANCE_FILE_BACKUP + DateTime.Now.Day.ToString + MonthName(DateTime.Now.Month) + DateTime.Now.Year.ToString + "-" + DateTime.Now.Hour.ToString + "_" + DateTime.Now.Minute.ToString + "_" + DateTime.Now.Second.ToString)
                File.Delete(BALANCE_FILE_BACKUP)
                File.Move(BALANCE_FILE, BALANCE_FILE_BACKUP)
            End If
         Catch ex As Exception
            Return 1
        End Try

        '----Open the value file for writing--------------------------------
        Try
            oWriter = New StreamWriter(BALANCE_FILE, False)
            oWriter.AutoFlush = True
        Catch ex As Exception
            Console.WriteLine("Error opening new backup file for writing: " + ex.Message)
            Return 2
        End Try


        Try
            oWriter.WriteLine(dBalance.ToString)
            oWriter.WriteLine(strBalanceNote)
            oWriter.WriteLine(dToMove.ToString)
            oWriter.WriteLine(DateTime.Now.ToLongDateString + " " + DateTime.Now.ToLongTimeString)

            oWriter.Close()
            strLastUpdated = DateTime.Now.ToLongDateString + " " + DateTime.Now.ToLongTimeString
        Catch ex As Exception
            Return 3
        End Try

        Return 0

    End Function

    Public Function GetValueMoved() As Double

        Dim bValidResponse As Boolean = False
        Dim dResult As Double = 0
        Dim strValueRead As String = ""

        Console.WriteLine()
        Console.WriteLine("Enter value moved.")
        Console.Write("> ")

        While Not bValidResponse
            strValueRead = Console.ReadLine()

            Try
                dResult = Double.Parse(strValueRead.Trim)
                bValidResponse = True
            Catch ex As Exception
                Console.WriteLine("Please enter a valid currency value.")
                Console.Write("> ")
            End Try
        End While

        Return dResult

    End Function


    Public Function GetValueSpent() As Double

        Dim bValidResponse As Boolean = False
        Dim dResult As Double = 0
        Dim strValueRead As String = ""

        Console.WriteLine()
        Console.WriteLine("Enter value spent.")
        Console.Write("> ")

        While Not bValidResponse
            strValueRead = Console.ReadLine()

            Try
                dResult = Double.Parse(strValueRead.Trim)
                bValidResponse = True
            Catch ex As Exception
                Console.WriteLine("Please enter a valid currency value.")
                Console.Write("> ")
            End Try
        End While

        Return dResult

    End Function

    Public Function GetNewBalanceNote() As String
        Dim bValidResponse As Boolean = False

        Dim strValueRead As String = ""

        Console.WriteLine()
        Console.WriteLine("Enter new balance note.")
        Console.Write("> ")

        While Not bValidResponse
            strValueRead = Console.ReadLine()

            bValidResponse = True

        End While

        Return strValueRead

    End Function

    Private Function GetShoppingList() As String

        SHOPPING_LIST_FILE = ConfigurationManager.AppSettings("Shopping_List_File")

        Dim oSB As StringBuilder = New StringBuilder()
        oSB.AppendLine("Current contents of the shopping list file at [" + SHOPPING_LIST_FILE + "] follow below")
        oSB.AppendLine("------------------------------------------------------------------------------")

        Try
            Using oReader As New StreamReader(SHOPPING_LIST_FILE)

                While Not oReader.EndOfStream
                    oSB.AppendLine(oReader.ReadLine)
                End While

                oReader.Close()
            End Using
        Catch ex As FileNotFoundException
            Console.WriteLine("Could not find shopping list file. Should be at " + SHOPPING_LIST_FILE)
        End Try

        oSB.AppendLine("------------------------------------------------------------------------------")

        Return oSB.ToString()

    End Function

    Private Function GetConstantsFromFile()
        Dim bSuccess As Boolean = True

        CONSTANTS_FILE = ConfigurationManager.AppSettings("TCB_Constants")

        Try
            Using oReader As New StreamReader(CONSTANTS_FILE)

                While Not oReader.EndOfStream

                    Dim fileLine = oReader.ReadLine
                    ProcessConstantsFileLine(fileLine.Trim)

                End While

                oReader.Close()
            End Using
        Catch ex As FileNotFoundException
            Console.WriteLine("Could not find constants file. Should be at " + CONSTANTS_FILE)
            bSuccess = False
        End Try

        If LUNCH_ONLY <> 0 And JOURNEY <> 0 And LUNCH_LAST_INCREASE <> "" And BUSES_LAST_INCREASE <> "" Then
            JOURNEYS_ONLY = JOURNEY * 2
            FULL_DAY_UPDATE = (JOURNEY * 2) + LUNCH_ONLY
        Else
            Console.WriteLine("Not all the constants were read in. Expecting ""Lunch"" , ""Journey"" , ""Journey Last Increase"" and ""Lunch Last Increase""")
            bSuccess = False
        End If


        Return bSuccess

    End Function

    Public Sub ProcessConstantsFileLine(ByVal constLine As String)

        If constLine.Equals("") OrElse constLine.Substring(0, 1) = "#" Then
            Return
        End If

        Dim iColonPos = constLine.IndexOf(":")

        If iColonPos < 0 Then
            Return
        End If

        Dim split() As String = constLine.Split(":")

        Select Case split(0).Trim
            Case "Lunch"
                LUNCH_ONLY = CDbl(split(1))
            Case "Journey"
                JOURNEY = CDbl(split(1))
            Case "Journey Last Increase"
                BUSES_LAST_INCREASE = split(1).Trim
            Case "Lunch Last Increase"
                LUNCH_LAST_INCREASE = split(1).Trim
            Case Else
                Console.WriteLine("Unknown constant name in constants file : " + split(0))
        End Select
    End Sub



End Module
