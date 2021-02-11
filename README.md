# AUTOMATE D’ENVOI DE MAILS AVEC PIECE JOINTE

## Le choix des outils

J'ai programmé l'automate en VB .NET, comme demandé, et donc avec Visual Studio 2019.
Il existe des méthodes existantes pour envoyer des mails en VB mais elles sont considérés comme obsolète (manque de protocole). 
J'ai donc opté pour un package externe nommé MailKit.

## Structure des dossiers

Il faut deux dossiers pour que le programme fonctionne, situés au même niveau que l'exécutable.
Le premier doit être nommé "A_traiter" et contiendra tous les dossiers qui permettront l'envoi d'un mail.
Le second doit être nommé "Backup" et contiendra les dossiers déjà traités.
N.B : Ces dossiers sont fournis dans le Git.

Pour qu'un mail soit envoyé, il faut mettre un dossier dans le répertoire A_traiter. 
Ce dossier peut avoir n'importe quel nom mais doit contenir obligatoirement un fichier .dst contenant au moins une adresse mail.
