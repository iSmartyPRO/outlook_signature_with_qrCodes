On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strName = objUser.description
strTitle = objUser.Title
strDepartment = objUser.Department
strCompany = objUser.Company
strDirectPhone = "+7 (812) 770 60 01"
strFax = "+7 (812) 770 60 02"
strEmail = objUser.mail
strAddress = Replace(objUser.streetAddress,vbCrLf, ", ") 'Changes mutli-line address to one line
strPostCode = objUser.postalCode
strCity = objUser.l
strState =  objUser.st
strCountry = objUser.c
if (objUser.mobile) Then strMobile = ", M: " & objUser.mobile else strMobile = "" End if
strSwitchPhone = objUser.otherTelephone
strSkype = objUser.ipPhone
strWebsite = objUser.wWWHomePage
strExt = " (" & objUser.pager & ")"
strLogo = "\\fs-genco.gencoindustry.com\scripts$\outlook_signature_with_qrCodes\images\logo.png"
strQrCode = "\\fs-genco.gencoindustry.com\scripts$\outlook_signature_with_qrCodes\images\qr_codes\qr_code_" & objUser.samAccountName & ".png"
strConfidential = "КОНФИДЕНЦИАЛЬНОСТЬ" & vbCrLf & "Настоящее электронное письмо и приложения к нему содержат информацию, составляющую коммерческую тайну." & vbCrLf & vbCrLf & "Указанная информация не может быть использована, скопирована или разглашена Вами, если согласие на выполнение таких действий ранее не было предоставлено Вам обладателем такой информации." & vbCrLf & vbCrLf & "Если Вы получили настоящее электронное письмо по ошибке либо Вам не был ранее предоставлен доступ к информации, содержащейся в настоящем электронном письме и приложениях к нему, пожалуйста, немедленно поставьте в известность отправителя и удалите данное электронное письмо и приложения к нему."
wdColorBlack = 0
wdColorBlue = 16711680
wdColorGray = 5855577
fontSize = 10
fontName = "Arial"
fontColor = RGB(144,140,140)
linkColor = RGB (000,045,154)

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
set objRange = objDoc.Range()

Set objTable = objDoc.Tables.Add(objRange,10,2)

'Универсальные свойства
For i = 1 to 10
    With objTable.Cell(i, 2)
        .width = 600
        .Range.ParagraphFormat.SpaceAfter = 0
        .Range.ParagraphFormat.SpaceAfterAuto = False
        .Range.ParagraphFormat.LineSpacingRule = 0
        .Range.Font.Size = fontSize
        .Range.Font.Name = fontName
        .Range.Font.Color = fontColor
    End With
Next

With objTable
    .Styles("Hyperlink").Font.Color = wdColorBlue
    .Paragraphs.SpaceAfter = 0
    .width = 100
    With .Cell(1,1)
        .Merge objTable.Cell(objTable.Rows.Count,1)
        .width = 100
        .Range.ParagraphFormat.Alignment = 1
        '.Range.Cells.VerticalAlignment = 1
        .Range.InlineShapes.AddPicture(strLogo)
    End With
    
    .Cell(1,2).Range.Text = "C уважением,"
    .Cell(2,2).Range.Text = strName: .Cell(2,2).Range.Font.Bold = True
    .Cell(3,2).Range.Text = strTitle & " • " & strDepartment: .Cell(3,2).Range.Font.Italic = True
    .Cell(4,2).Range.Text = strCity & ", " & strAddress
    .Cell(5,2).Range.Text = "Т:" & strDirectPhone & strExt & ", " & "Ф: " & strFax & ", " & "М:" & strMobile
    .Cell(6,2).Range.Text = strEmail & ", www.gencoindustry.com"
    .Cell(6,2).Range.Font.Color = linkColor
    .Cell(8,2).width = 170
    .Cell(8,2).Range.InlineShapes.AddPicture(strQrCode)
    .Cell(8,2).width = 600
    .Cell(10,2).Range.Text = strConfidential
    .Cell(10,2).Range.Font.Italic = True
End With

objSignatureEntries.Add "GENCO - New", objDoc.Range()
objSignatureObject.NewMessageSignature = "GENCO - New"
objDoc.Saved = True
objWord.Quit


Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
objSelection.Font.Name = "Arial"
objSelection.Font.Size = 9
objSelection.Font.Color = RGB(144,140,140) 
objSelection.TypeParagraph()
objSelection.TypeText "C уважением,"
objSelection.TypeText Chr(11)
objSelection.Font.Bold = true
objSelection.TypeText strName
objSelection.Font.Bold = false
objSelection.TypeText Chr(11)
objSelection.TypeText strtitle & ", " & strDepartment
Set objSelection = objDoc.Range()
objSignatureEntries.Add "GENCO - Reply", objSelection
objSignatureObject.ReplyMessageSignature = "GENCO - Reply"
objDoc.Saved = True
objWord.Quit




Set curSelection = Nothing
Set objLink = Nothing
Set objWord = Nothing
Set objDoc = Nothing
Set objSelection = Nothing
Set objSignatureObject = Nothing
Set objSignatureEntries = Nothing
