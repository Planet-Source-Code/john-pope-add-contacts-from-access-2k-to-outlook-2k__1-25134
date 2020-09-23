<div align="center">

## Add Contacts from Access 2k to outlook 2k


</div>

### Description

This code will send customer information such as Name address , city, state, zip to outlook2 contacts
 
### More Info
 
You will need to set the references to use the Outlook 9.0 Object library.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Pope](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-pope.md)
**Level**          |Intermediate
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VBA MS Access
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-pope-add-contacts-from-access-2k-to-outlook-2k__1-25134/archive/master.zip)





### Source Code

```
Private Sub cmdAddOutlook_Click()
  'I have used this code in two ways. I have made this one here just work in the cmdbutton that I created on the Access form and placed all the code in here
  'You can also make this a function
  Dim oOutlook As Outlook.Application
  Dim oContact As Outlook.ContactItem
  'Create Object
  Set oOutlook = New Outlook.Application
  'Create and new Contact
  Set oContact = oOutlook.CreateItem(olContactItem)
  With oContact
    'Change what you need in here. All Vairables after the = sign are fields in my database.
    'You will need to change them to fields that you have in yours.
    'To find a list of the items here you can set go to the object browser.
    .FullName = LastName & "," & FirstName
    .BusinessAddress = Address
    .BusinessAddressCity = City
    .BusinessAddressState = State
    .BusinessAddressPostalCode = Zip
    .HomeTelephoneNumber = HomePhone
    .BusinessTelephoneNumber = BusPhone
    .MobileTelephoneNumber = CellPhone
    .Email1Address = Email
    .CompanyName = CompanyName
    .Categories = CompanyName
    .Save
    End With
  'Change msgbox ot what you want it to say. I just left it simple
  MsgBox "Contact has been Added", vbInformation
  'Release Outlook
  Set oOutlook = Nothing
End Sub
```

