## Attendance Check
Genrererar en fin excel fil automatiskt från närvarolistor som hämtas från teams.

### Vart man lägger närvaro listorna
Närvarofilerna (meetingAttendanceList.csv) som laddas ner från teams under mötets gång läggs inne i 'PutAllAttendanceCsvsHere' foldern.
Man behöver aldrig ta bort några filer, bara lägga in nya och sen generera om excelen.

### Hur man ställer in sina klasser
* Filen måste heta 'StudentList.txt' och ligga i 'AttendanceCheck' projektet
* Alla elevgrupper måste ha ett namn inom klamrar.
* Det måste finnas minst en grupp.
* Elevnamnen i denna filen måste matcha namnen inne på teams.
* Om en elev hoppar av är det bara att kommentera ut med # eller att ta bort raden.

StudentList.txt -- exempel
```
# kommentarer med hashtag

[Ystad]
Börje Svensson
Cammilla Jansson
```

### Guide för att hämta namn via PowerShell
 source: 
 * https://answers.microsoft.com/en-us/education_ms/forum/edu_msteams-edu_teams/export-members-list-on-teams/1dbfe04d-d388-4ed4-9183-64e21904c380
 * https://answers.microsoft.com/en-us/msoffice/forum/msoffice_o365admin-mso_teams-mso_o365b/the-term-connect-microsoftteams-is-not-recognized/292044e2-b841-48d9-a902-57ccbb10e64e
 * https://techcommunity.microsoft.com/t5/microsoft-teams/get-team-user-powershell-for-teams/m-p/499316
 
Kör powershell som admin
 ```
 >Find-Module MicrosoftTeams
 >Install-Module MicrosoftTeams
 >Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
 >Connect-MicrosoftTeams
 >get-team -User [samma mail som du loggade in med]
 >Get-TeamUser -GroupId [team hash] -Role Member
```
Alt+Mus för att blockkopiera namnen in hit sen
