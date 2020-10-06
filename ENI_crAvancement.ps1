##[Ps1 To Exe]
##
##NcDBCIWOCzWE8pGP3wFk4Fn9fksqYsyehZKi14qo8PrQlDDPW5sTTGh4kj2uEFOpF/cKUJU=
##Kd3HDZOFADWE8uK1
##Nc3NCtDXThU=
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiS5
##OsHQCZGeTiiZ4tI=
##OcrLFtDXTiW5
##LM/BD5WYTiiZ4tI=
##McvWDJ+OTiiZ4tI=
##OMvOC56PFnzN8u+VslQ=
##M9jHFoeYB2Hc8u+VslQ=
##PdrWFpmIG2HcofKIo2QX
##OMfRFJyLFzWE8uK1
##KsfMAp/KUzWI0g==
##OsfOAYaPHGbQvbyVvnQnqxqO
##LNzNAIWJGmPcoKHc7Do3uAu+DDlL
##LNzNAIWJGnvYv7eVvnQX
##M9zLA5mED3nfu77Q7TV64AuzAgg=
##NcDWAYKED3nfu77Q7TV64AuzAgg=
##OMvRB4KDHmHQvbyVvnQX
##P8HPFJGEFzWE8tI=
##KNzDAJWHD2fS8u+Vgw==
##P8HSHYKDCX3N8u+VwDV77E7gRmdrRNCXsLOppA==
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTlaDjofG5iZk2WjvTmEuUuGeqr2zy5GA0P/6qSTeTKYeXFh+k2f5HE7d
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
# Lecture du fichier de paramètres
Get-Content params.txt | Foreach-Object{
   $var = $_.Split('=')
   New-Variable -Name $var[0] -Value $var[1]
}

# On récupère les fichiers .odt dans le répertoire courant
$chapitres = Get-ChildItem -Path . -Name -Recurse -Include *.odt

# Compteurs
$caracteresTotal = 0
$chapitreTotal = 0
$listeChapitre = ""

# On compte les caractères de chaque chapitre (fichier odt)
ForEach ($c in $chapitres) {
    $chapitreTotal++
    $nbreDeCaracteresDuChapitre = .\odt2txt.exe $c | Measure-Object -character | Select-Object -ExpandProperty Characters
    $listeChapitre += "$chapitreTotal : $nbreDeCaracteresDuChapitre caracteres"
    $caracteresTotal += $nbreDeCaracteresDuChapitre
}

# On prépare le mail
$body = "Bonjour,`r`n"
$body += "Comme tous les mois, un point sur l'avancement de mon ouvrage : $titre`r`n`r`n`r`n"
$body += "=============Par chapitre=============`r`n"
Foreach ($l in $listeChapitre) {
    $body += "$l `r`n"
}
$body += "`r`n===`r`n"
$body += "Nombre de chapitres : $chapitreTotal `n"
$body += "Nombre de caracteres au total : $caracteresTotal"
if ($objectifCaracteres -ne $null) {
    $pourcentage = $caracteresTotal/$objectifCaracteres*100
    $body += "/$objectifCaracteres soit $pourcentage%"
}
$body += "`r`n"

# On demande le login/mdp de messagerie
$credential = Get-Credential

# On définit les paramètres de messagerie
$paramEmail = @{
    SmtpServer                 = $smtp
    Port                       = $port
    UseSSL                     = $true
    Credential                 = $credential
    From                       = $expediteur
    To                         = $destinataire
    Subject                    = "Avancement ouvrage $titre"
    Body                       = $body
    DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
}

# On envoie le mail
if ($mail -eq $true) {
    Send-MailMessage @paramEmail
}