 Set loc = CreateObject("WbemScripting.SWbemLocator")
 Dim WbemServices
 Set WbemServices = loc.ConnectServer("bmobccsccmc01", "root\sms\site_B01")
    Set clsScheduleMethods = WbemServices.Get("SMS_ScheduleMethods")

    sInterval = "0078AB4000080000"
    clsScheduleMethods.ReadFromString sInterval, avTokens
    For each vToken In avTokens
        wscript.echo vToken.GetObjectText_
    Next
