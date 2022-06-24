Module Vars
    Public Const arrDivider = ","
    '---defines whether game is clasic or samurai variant---
    Public blnSamurai As Boolean
    Public intBGTransparency As Double = 0.23
    Public intBorderTransparency As Double = 0.7
    Public intMO_BGTransparency As Double = 0.35
    Public intM0_BorderTransparency As Double = 1.0
    '---for batch solving---
    Private StopThreadWorkFlag As Boolean
    Public StopThreadWork As Boolean
    '---for walkthrough---
    Public StopShowAllSteps As Boolean
    Public blnAllSteps As Boolean
End Module
