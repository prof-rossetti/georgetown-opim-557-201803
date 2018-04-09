'
' DISPLAYS MODULE
' ... Author: Prof. Rossetti <prof.mj.rossetti@gmail.com>.
' ... License: Students, feel free but not obligated to use this module in your project as long as you retain this attribution section. If you wrote something like this on your own, no need to attribute. If this code inspired you to write your own code, please still consider providing an attribution link to this file's GitHub URL.
'

' LogUserInputs displays a message box with nicely-formatted user input values.
Public Sub LogUserInputs(ByVal Age As Integer, ByVal RetirementAge As Integer, ByVal SavingsBalance As Double, ByVal AnnualContribution As Double, ByVal AnnualInterestRate As Double)
    MsgBox ("INFORMATION INPUTS" & vbNewLine & _
            "---------------------------------" & vbNewLine & _
            "Current Age: " & Age & vbNewLine & _
            "Retirement Age: " & RetirementAge & vbNewLine & _
            "Savings Balance: " & FormatUSD(SavingsBalance) & vbNewLine & _
            "Annual Contribution: " & FormatUSD(AnnualContribution) & vbNewLine & _
            "Annual Interest Rate: " & FormatPct(AnnualInterestRate) _
    )
End Sub

' LogFinalOutputs displays a message box with nicely-formatted final output values.
Public Sub LogFinalOutputs(ByVal SavingsBalance As Double, ByVal TotalContribution As Double, ByVal TotalInterest As Double)
    Dim PctContribution As Double
    Dim PctInterest As Double

    PctContribution = TotalContribution / SavingsBalance
    PctInterest = TotalInterest / SavingsBalance

    MsgBox ("INFORMATION OUTPUTS" & vbNewLine & _
            "---------------------------------" & vbNewLine & _
            "Retirement Savings Balance: " & FormatUSD(SavingsBalance) & vbNewLine & _
            "Total Contribution: " & FormatUSD(TotalContribution) & " (" & FormatPct(PctContribution) & ")" & vbNewLine & _
            "Total Interest Accrued: " & FormatUSD(TotalInterest) & " (" & FormatPct(PctInterest) & ")" & vbNewLine _
    )
End Sub
