Randomize ' Initialize random number generator
Dim targetNumber, userGuess, attempts, maxAttempts
targetNumber = Int((100 - 1 + 1) * Rnd + 1) ' Generate a random number between 1 and 100
attempts = 0
maxAttempts = 25 ' Set maximum number of attempts

Do
    userGuess = InputBox("Guess a number between 1 and 100 (Attempt " & (attempts + 1) & " of " & maxAttempts & "):", "Guess It!")
    
    ' Check if user clicked Cancel
    If userGuess = "" Then
        MsgBox "Game cancelled. The correct number was " & targetNumber, vbInformation, "Guess It!"
        WScript.Quit
    End If

    ' Convert input to a number
    userGuess = CInt(userGuess)
    attempts = attempts + 1

    If userGuess < targetNumber Then
        MsgBox "Too low! Try again.", vbExclamation, "Guess It!"
    ElseIf userGuess > targetNumber Then
        MsgBox "Too high! Try again.", vbExclamation, "Guess It!"
    Else
        MsgBox "Congratulations! You guessed the correct number in " & attempts & " attempts.", vbInformation, "Guess It!"
        Exit Do
    End If

    ' Check if maximum attempts reached
    If attempts >= maxAttempts Then
        MsgBox "Sorry, you've reached the maximum number of attempts. The correct number was " & targetNumber, vbCritical, "Guess It!"
        Exit Do
    End If
Loop
