
Questions:  
1. MCHS, data ready flag keeps at 0. (how to repeat the fault? power cycle solve the problem.)  




[threading read write](https://stackoverflow.com/questions/2439122/problem-with-two-net-threads-and-hardware-access)  
[Threading documentation from Microsoft](https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/concepts/threading/)  


very important --- how to convert between hex and numerical
[hex to num](https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/types/how-to-convert-between-hexadecimal-strings-and-numeric-types)  

[byte array concat, .CopyTo](https://stackoverflow.com/questions/1547252/how-do-i-concatenate-two-arrays-in-c)  



    Public Const CMD_PRODUCTION_LOGIN As String = "734D49000006E80C744D"
    Public Const RES_PRODUCTION_LOGIN_LENGTH As Integer = 15


    Public Const CMD_MCAVERAGINGDEPTH As String = "73574E204D43414420"
    Public Const RES_MCAVERAGINGDEPTH_LENGTH As Integer = 18

    Public Const CMD_CLEAR_STATISTICOUTDONE_FLAG As String = "73574E204D4348532000"
    Public Const RES_CLEAR_STATISTICOUTDONE_FLAG_LENGTH As Integer = 18

    Public Const CMD_READ__STATISTICOUTDONE_FLAG As String = "73524E204D434853"
    Public Const RES_READ__STATISTICOUTDONE_FLAG_LENGTH As Integer = 19

    Public Const CMD_MCDISTRSSISTATISTIC As String = "73524E204D435354"
    Public Const RES_MCDISTRSSISTATISTIC_LENGTH As Integer = 62



https://stackoverflow.com/questions/4130194/what-is-the-difference-between-task-and-thread  
to study difference between task and thread.

[task repeat](https://stackoverflow.com/questions/7472013/how-to-create-a-thread-task-with-a-continuous-loop)

how to solve the int conversion problem.
Answer: use bitwise operation. Pay attention to the byte size differences.

[text change](https://msdn.microsoft.com/en-us/library/system.windows.forms.control.textchanged(v=vs.110).aspx)
