BeforeAll {
    #Lade...
    ."$PSScriptRoot\generate_data.ps1"
}
Describe "Generate-DummyData"{
    It "name des tests" {
        $result=Generate-DummyData
        $result.count | should -be [int] 
    }
}    