
Sub ReplaceTextAcrossOpenDocs()
Dim aFh(49) As String
aFh(1) = "Blackpool"
aFh(2) = "Bristol R."
aFh(3) = "Burton A."
aFh(4) = "Coventry"
aFh(5) = "Oxford Utd."
aFh(6) = "Peterboro"
aFh(7) = "Rochdale"
aFh(8) = "Southend"
aFh(9) = "Carlisle"
aFh(10) = "Cheltenham"
aFh(11) = "Crawley"
aFh(12) = "Exeter"
aFh(13) = "Leyton O."
aFh(14) = "Macclesfield"
aFh(15) = "Mansfield"
aFh(16) = "Morecambe"
aFh(17) = "Salford C."
aFh(18) = "Scunthorpe"
aFh(19) = "Stevenage"
aFh(20) = "Swindon"
aFh(21) = "Aldershot"
aFh(22) = "Barrow"
aFh(23) = "Eastleigh"
aFh(24) = "Ebbsfleet"
aFh(25) = "Fylde"
aFh(26) = "Halifax"
aFh(27) = "Notts Co."
aFh(28) = "Solihull M."
aFh(29) = "Stockport"
aFh(30) = "Wrexham"
aFh(31) = "Yeovil"
aFh(32) = "Darlington"
aFh(33) = "Belarus"
aFh(34) = "Bosnia"
aFh(35) = "Cyprus"
aFh(36) = "Denmark"
aFh(37) = "Estonia"
aFh(38) = "F. Islands"
aFh(39) = "Georgia"
aFh(40) = "Hungary"
aFh(41) = "Italy"
aFh(42) = "Kazakhstan"
aFh(43) = "Liechtenstein"
aFh(44) = "Malta"
aFh(45) = "Norway"
aFh(46) = "Poland"
aFh(47) = "Scotland"
aFh(48) = "Slovenia"
aFh(49) = "Wales"
    Dim aFA(49) As String
aFA(1) = "Rotherham"
aFA(2) = "Milton K.D."
aFA(3) = "Bolton"
aFA(4) = "Tranmere"
aFA(5) = "Doncaster"
aFA(6) = "Lincoln"
aFA(7) = "Accrington"
aFA(8) = "Wimbledon"
aFA(9) = "Crewe"
aFA(10) = "Newport Co."
aFA(11) = "Colchester"
aFA(12) = "Forest G."
aFA(13) = "Walsall"
aFA(14) = "Port Vale"
aFA(15) = "Oldham"
aFA(16) = "Bradford C."
aFA(17) = "Cambridge U."
aFA(18) = "Northampton"
aFA(19) = "Grimsby"
aFA(20) = "Plymouth"
aFA(21) = "Hartlepool"
aFA(22) = "Dover"
aFA(23) = "Chorley"
aFA(24) = "Maidenhead"
aFA(25) = "Sutton Utd."
aFA(26) = "Boreham W."
aFA(27) = "Torquay"
aFA(28) = "Bromley"
aFA(29) = "Dagenham"
aFA(30) = "Chesterfield"
aFA(31) = "Harrogate"
aFA(32) = "Boston"
aFA(33) = "Netherlands"
aFA(34) = "Finland"
aFA(35) = "Russia"
aFA(36) = "Switzerland"
aFA(37) = "Germany"
aFA(38) = "Romania"
aFA(39) = "R. Ireland"
aFA(40) = "Azerbaijan"
aFA(41) = "Greece"
aFA(42) = "Belgium"
aFA(43) = "Armenia"
aFA(44) = "Sweden"
aFA(45) = "Spain"
aFA(46) = "Macedonia"
aFA(47) = "San Marino"
aFA(48) = "Austria"
aFA(49) = "Croatia"
    Dim aN(49) As Integer
    For i = 1 To 49
    aN(i) = i
    Next i
    

    Const txtFIND As String = "Swidon"
    Const txtREPLACE As String = "Man Utd"
    
    Dim d As Document
    Dim p As Page
    
    For Each d In Documents
    For Each p In d.Pages
    p.TextReplace txtFIND, txtREPLACE, True, False
    Next p
    Next d
End Sub
