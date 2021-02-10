Imports System.Text
Imports System.IO


Module main
    Const nbDestinataires As Integer = 10 'Les nombres max de destinataires et de pi�ces jointes pour chaque mail
    Const nbPiecesJointes As Integer = 10

    Function get_pieces_jointes(chemin As String, nom_fichier As String) 'Renvoie la liste des pi�ces jointes du fichier .dst (ligne commencant par FPG_)
        Dim liste_pieces_jointes(nbPiecesJointes) As String
        Dim index_pj As Integer = 0
        Dim flux_lecture As StreamReader = New StreamReader(chemin + nom_fichier)
        Dim regex_pj As String = "^(FPG_.+)$"
        Dim ligne As String
        Do
            ligne = flux_lecture.ReadLine()
            If ligne IsNot Nothing Then
                If RegularExpressions.Regex.IsMatch(ligne, regex_pj) And ligne.Length > 4 Then  'Pour pouvoir enlever FPG_ sans avoir un nom vide
                    ligne = ligne.Substring(4)
                    liste_pieces_jointes(index_pj) = ligne + ".pdf"
                    index_pj += 1
                    Console.WriteLine(ligne)
                End If
            End If
        Loop Until ligne Is Nothing
        Return liste_pieces_jointes
    End Function

    Function get_receveurs(chemin As String, nom_fichier As String) 'Renvoie la liste des receveurs (destinataires) du fichier .dst (les adresses mails)
        Dim liste_receveurs(nbDestinataires) As MimeKit.MailboxAddress
        Dim index_receveurs As Integer = 0
        Dim flux_lecture As StreamReader = New StreamReader(chemin + nom_fichier)
        Dim regex_mail As String = "^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$" 'regex d'une adresse mail trouv� sur Internet
        Dim ligne As String
        Do
            ligne = flux_lecture.ReadLine()
            If ligne IsNot Nothing Then
                If RegularExpressions.Regex.IsMatch(ligne, regex_mail) Then
                    liste_receveurs(index_receveurs) = New MimeKit.MailboxAddress(ligne)
                    index_receveurs += 1
                End If
            End If
        Loop Until ligne Is Nothing
        Return liste_receveurs
    End Function

    Function get_texte_mail(chemin As String, nom_fichier As String) 'Renvoie le contenu du fichier texte fourni (corps du mail)
        Dim texte As String = File.ReadAllText(chemin + nom_fichier, Encoding.UTF8)
        Return texte
    End Function

    Sub envoi_mail(chemin As String)

        Dim message As New MimeKit.MimeMessage() 'Mail dans sa globalit�



        Dim fichier_dst As String = Dir(chemin + "*.dst") 'On prend le premier fichier .dst trouv�
        Dim fichier_txt As String = Dir(chemin + "*.txt") 'On prend le premier fichier .txt trouv�

        Dim liste_pieces_jointes() As String = get_pieces_jointes(chemin, fichier_dst)
        Dim liste_receveurs() As MimeKit.MailboxAddress = get_receveurs(chemin, fichier_dst)

        message.From.Add(New MimeKit.MailboxAddress("Ruben Delamarche", "rub.del@hotmail.fr")) 'L'envoyeur

        For Each addr_receveur As MimeKit.MailboxAddress In liste_receveurs
            If (addr_receveur IsNot Nothing) Then
                message.To.Add(addr_receveur)
            End If
        Next



        message.Subject = "Message du bot Camtek"
        Dim data As New MimeKit.BodyBuilder() 'Outil qui va contenir le corps du mail (texte) et les pieces jointes
        Dim texte_mail As String = get_texte_mail(chemin, fichier_txt)
        data.TextBody = texte_mail

        For Each nom_pj As String In liste_pieces_jointes
            If (nom_pj IsNot Nothing) Then
                data.Attachments.Add(chemin + nom_pj)
                Console.WriteLine(chemin + nom_pj)
            End If
        Next

        message.Body = data.ToMessageBody()


        Dim client As New MailKit.Net.Smtp.SmtpClient()
        client.Connect("smtp.gmail.com", 465, True)
        client.Authenticate(Encoding.UTF8, "neventus2.0@gmail.com", "fafaliba")
        client.Send(message)
        client.Disconnect(True)

    End Sub

    Sub Main(args As String())
        Dim dossier_src As New IO.DirectoryInfo("Source\")

        'Rajouter un while pour faire une boucle infinie

        For Each sous_dossier In dossier_src.GetDirectories
            envoi_mail("Source\" + sous_dossier.Name + "\")
        Next


    End Sub


End Module

