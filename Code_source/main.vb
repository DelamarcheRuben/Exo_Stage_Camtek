Imports System.Text
Imports System.IO

Module main
    Const nbDestinataires As Integer = 10 'Les nombres max de destinataires et de pieces jointes pour chaque mail
    Const nbPiecesJointes As Integer = 10
    Const nom_dossier_a_traiter As String = "A_traiter\"
    Const nom_dossier_backup As String = "Backup\"
    Const duree_attente_ms As Integer = 5000

    Function get_pieces_jointes(chemin As String, nom_fichier As String) 'Renvoie la liste des pieces jointes du fichier .dst (ligne commencant par FPG_)
        Dim liste_pieces_jointes(nbPiecesJointes - 1) As String
        Dim index_pj As Integer = 0
        Dim flux_lecture As StreamReader = New StreamReader(chemin + nom_fichier)
        Dim regex_pj As String = "^(FPG_.+)$" 'tout fichier commencant par FPG_ soit une piece jointe
        Dim ligne As String
        Do
            ligne = flux_lecture.ReadLine()
            If ligne IsNot Nothing Then
                If RegularExpressions.Regex.IsMatch(ligne, regex_pj) And ligne.Length > 4 Then  'Pour pouvoir enlever FPG_ sans avoir un nom vide apres
                    ligne = ligne.Substring(4) 'on enleve le FPG_
                    liste_pieces_jointes(index_pj) = ligne + ".pdf"
                    index_pj += 1
                End If
            End If
        Loop Until ligne Is Nothing Or index_pj >= nbPiecesJointes - 1 'Si fichier fini ou max pj atteint
        flux_lecture.Close()
        Return liste_pieces_jointes
    End Function

    Function get_receveurs(chemin As String, nom_fichier As String) 'Renvoie la liste des receveurs (destinataires) du fichier .dst (les adresses mails)
        Dim liste_receveurs(nbDestinataires - 1) As MimeKit.MailboxAddress
        Dim index_receveurs As Integer = 0
        Dim flux_lecture As StreamReader = New StreamReader(chemin + nom_fichier)
        Dim regex_mail As String = "^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$" 'regex d'une adresse mail trouve sur Internet
        Dim ligne As String
        Do
            ligne = flux_lecture.ReadLine()
            If ligne IsNot Nothing Then
                If RegularExpressions.Regex.IsMatch(ligne, regex_mail) Then
                    liste_receveurs(index_receveurs) = New MimeKit.MailboxAddress(ligne)
                    index_receveurs += 1
                End If
            End If
        Loop Until ligne Is Nothing Or index_receveurs >= nbDestinataires - 1 'Si fichier fini ou max destinataires atteint
        flux_lecture.Close()
        Return liste_receveurs
    End Function

    Function get_texte_mail(chemin As String, nom_fichier As String) 'Renvoie le contenu du fichier texte fourni (corps du mail)
        Dim texte As String
        If nom_fichier <> "" Then
            texte = File.ReadAllText(chemin + nom_fichier, Encoding.UTF8)
        Else texte = ""
        End If

        Return texte
    End Function

    Sub envoi_mail(chemin As String)

        Dim message As New MimeKit.MimeMessage() 'Mail dans sa globalite

        Dim fichier_dst As String = Dir(chemin + "*.dst") 'On prend le premier fichier .dst trouve
        Dim fichier_txt As String = Dir(chemin + "*.txt") 'On prend le premier fichier .txt trouve

        If fichier_dst = "" Then
            Console.WriteLine("Fichier dst manquant dans : " + chemin)
            Exit Sub
        End If

        Dim liste_pieces_jointes() As String = get_pieces_jointes(chemin, fichier_dst)
        Dim liste_receveurs() As MimeKit.MailboxAddress = get_receveurs(chemin, fichier_dst)

        If liste_receveurs.ElementAt(0) Is Nothing Then 'on peut envoyer un mail vide ou sans pj mais pas sans destinataires
            Console.WriteLine("Destinataires manquants : " + chemin)
            Exit Sub 'On quitte le traitement du mail si on en a pas
        End If

        message.From.Add(New MimeKit.MailboxAddress("Ruben Delamarche", "ruben.delamarche@universite-paris-saclay.fr")) 'L'envoyeur, peut etre changer

        For Each addr_receveur As MimeKit.MailboxAddress In liste_receveurs
            If (addr_receveur IsNot Nothing) Then
                message.To.Add(addr_receveur)
            End If
        Next



        message.Subject = "Message du bot Camtek"
        Dim data As New MimeKit.BodyBuilder() 'Outil qui va contenir le corps du mail (texte) et les pieces jointes
        Dim texte_mail As String = get_texte_mail(chemin, fichier_txt)
        data.TextBody = texte_mail 'si le fichier txt n'existait pas ou etait vide alors le mail aura un corps vide mais sera envoye quand meme

        For Each nom_pj As String In liste_pieces_jointes
            If (nom_pj IsNot Nothing) Then
                Try
                    data.Attachments.Add(chemin + nom_pj)
                Catch ex As Exception
                    Console.WriteLine("Pièce jointe manquante : " + chemin + nom_pj)
                End Try

            End If
        Next

        message.Body = data.ToMessageBody()

        Try
            Dim client As New MailKit.Net.Smtp.SmtpClient()
            client.Connect("smtp.gmail.com", 465, True)
            client.Authenticate(Encoding.UTF8, "exo.mail.stage@gmail.com", "jEHmTBj9S7S4Gzt") 'Adresse mail créée pour l'occasion
            client.Send(message)
            client.Disconnect(True)
        Catch ex As Exception 'Si jamais l'erreur se passe dans l'envoi (probleme mail, port, ...)
            Console.WriteLine("Problème dans l'envoi du mail : " + chemin)
        End Try


    End Sub

    Sub Main(args As String())
        Dim dossier_src As New IO.DirectoryInfo(nom_dossier_a_traiter) 'Pour ensuite récupérer tous les sous-dossiers
        Console.WriteLine("Traitement en cours...")
        While True 'On fait le traitement en boucle
            For Each sous_dossier In dossier_src.GetDirectories
                envoi_mail(nom_dossier_a_traiter + sous_dossier.Name + "\")
                Try
                    Directory.Move(nom_dossier_a_traiter + sous_dossier.Name, nom_dossier_backup + sous_dossier.Name) 'On tente de le déplacer dans le dossier backup
                Catch ex As Exception
                    Directory.Delete(nom_dossier_a_traiter + sous_dossier.Name, True) 'si probleme (repertoire deja exisstant par ex) alors on le supprime juste lui et ses fichiers
                End Try

            Next
            Threading.Thread.Sleep(duree_attente_ms) 'L'application analyse les dossiers à traiter toutes les 5 secondes (attention si il y a 2 dossiers 
            'à traiter alors ils seront traités sans écart de temps
        End While




    End Sub


End Module

