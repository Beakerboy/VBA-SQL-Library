Attribute VB_Name = "SQLUnitTests"
Public Function RunTests()
    Dim TestConfig As iTestableProject
    Dim SQLLibTestConfig As New SQLlibTests
    Set TestConfig = SQLLibTestConfig
    TestConfig.Run
End Function
