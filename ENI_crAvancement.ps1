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
    $listeChapitre += "$chapitreTotal : $nbreDeCaracteresDuChapitre caracteres `r`n"
    $caracteresTotal += $nbreDeCaracteresDuChapitre
}

# On prépare le mail
$body = "Bonjour,`r`n"
$body += "Comme tous les mois, un point sur l'avancement de mon ouvrage : $titre`r`n`r`n`r`n"
$body += "=============Par chapitre=============`r`n"
Foreach ($l in $listeChapitre) {
    $body += "$l"
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
