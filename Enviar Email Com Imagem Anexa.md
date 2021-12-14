Sub createJpg(Namesheet As String, nameRange As String, nameFile As String)
    ThisWorkbook.Activate
    Worksheets("DISPONIBILIDADE_ARU").Activate
    Set Plage = ThisWorkbook.Worksheets("DISPONIBILIDADE_ARU").Range("A1:G19")
    Plage.CopyPicture
    With ThisWorkbook.Worksheets("DISPONIBILIDADE_ARU").ChartObjects.Add(Plage.Left, Plage.Top, Plage.width, Plage.height)
        .Activate
        .Chart.Paste
        .Chart.Export Environ$("temp") & "\" & nameFile & ".jpg", "JPG"
    End With
    Worksheets("DISPONIBILIDADE_ARU").ChartObjects(Worksheets("DISPONIBILIDADE_ARU").ChartObjects.Count).Delete
Set Plage = Nothing
End Sub

Sub DISPONIBILIDADE_ARU_Thiago()
        Application.Calculation = xlManual
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
        End With
    
    Dim DestTo As String
    Dim DestCC As String
        DestTo = Sheets("DISPONIBILIDADE_ARU").Range("D20").Value 'Emails para (alterar essa célula se vc mudar o lugar)
        DestCC = Sheets("DISPONIBILIDADE_ARU").Range("D25").Value 'Emails em cópia (alterar essa célula se vc mudar o lugar)
     
        Dim TempFilePath As String
         
        'Create a new Microsoft Outlook session
        Set appOutlook = CreateObject("outlook.application")
        'create a new message
        Set Message = appOutlook.CreateItem(olMailItem)


​         
        With Message
            .Subject = "Disponibilidade ARU"
     
            .HTMLBody = "<span LANG=EN>" _
                & "<p class=style2><span LANG=EN><font FACE=Calibri SIZE=3>" _
                & "boa tarde,<br ><br >Segue Disponibilidade!" _


​                
            'first we create the image as a JPG file
            Call createJpg("Dashboard", "A1:G19", "DashboardFile")
            'we attached the embedded image with a Position at 0 (makes the attachment hidden)
            TempFilePath = Environ$("temp") & "\"
            .Attachments.Add TempFilePath & "DashboardFile.jpg", olByValue, 0
                
            'Then we add an html <img src=''> link to this image
            'Note than you can customize width and height - not mandatory
                
            .HTMLBody = .HTMLBody & "<br><B></B><br>" _
                & "<br>" _
                & "<img src='cid:DashboardFile.jpg'" & "width='500'height='350'<br>" _
                & "<br><br>Gestão Operacional,<br>" _
                & "Atc,</font></span>"
    
             .To = DestTo
             .Cc = DestCC
                               
            .Display
            .Send
        End With
     
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
        End With
        Application.Calculation = xlCalculationAutomatic
    End Sub