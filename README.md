#AUTOMATE D’ENVOI DE MAILS AVEC PIECE JOINTE

##Le choix des outils

J'ai programmé l'automate en VB .NET, comme demandé, et donc avec Visual Studio 2019.
Il existe des méthodes existantes pour envoyer des mails en VB mais elles sont considérés comme obsolète (manque de protocole). 
J'ai donc opté pour un package externe nommé MailKit. Pour envoyer des mails, j'ai utilisé un serveur SMTP de gmail. J'ai donc créé une adresse gmail qui enverra les mails. Attention, dans le code source le mot de passe est en clair.

##Structure des dossiers

Il faut deux dossiers pour que le programme fonctionne, situés au même niveau que l'exécutable.
Le premier doit être nommé "A_traiter" et contiendra tous les dossiers qui permettront l'envoi d'un mail.
Le second doit être nommé "Backup" et contiendra les dossiers déjà traités.
N.B : Ces dossiers sont fournis dans le Git.

Pour qu'un mail soit envoyé, il faut mettre un dossier dans le répertoire A_traiter. 
Ce dossier peut avoir n'importe quel nom et doit contenir un fichier .dst et un fichier .txt.
Le fichier dst contient:
  * Les noms des pièces jointes préfixés de FPG_ (ex : pj.pdf devient FPG_pj dans le fichier). Ces noms sont facultatifs.
  * Les adresses mail des destinataires. Au moins une adresse est nécessaire
Le fichier txt contient le corps du mail. Il est facultatif.
Attention si le dossier contient plusieurs ficheirs .txt ou .dst alors un seul sera pris en compte.

##Fonctionnement

Le programme tourne en boucle et à intervalle régulier analyse le dossier A_traiter et  traite l'intégralité des dossiers présents dedans.
Cette intervalle est de 5 secondes pour le moment. Ce nombre peut être modifié dans le code.

Chaque mail envoyé peut contenir par défaut 10 destinataires et 10 pièces jointes. Ce nombre peut être modifié dans le code.


##Point d'amélioration

* La taille de l'exécutable est lourd
* Au lancement du programme proposer des paramètres à l'utilisateur (ex: nombres de destinataires maximum ou fréquence de traitement)
* Crypter le mot de passe de l'adresse mail...
* ... ou utiliser un serveur SMTP à installer sur la machine directement (plus compliqué)



##Charge de travail

Temps de codage avec recherche des solutions et mise en place = 2h-2h30
Solution pour le serveur SMTP = 1h30
Création de l'exécutable exportable = 2h

Mon plus gros problème a été la création de l'exécutable avec Visual Studio.
Le problème venait d'une erreur de compréhension et du blocage de l'adresse gmail sur un ordinateur non vérifié.
