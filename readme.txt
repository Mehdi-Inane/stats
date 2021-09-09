--Ce programme utilise les modules Openpyxl, Matplotlib, scipy,  numpy, tkinter de Python 3.8.0--

--Mode de compilation du code--
Pour une compilation normale, se placer dans le dossier du programme et taper sur le terminal python3 stats.py
Pour générer un éxecutable transférable sur les différentes machines, il faudra d'abord installer pyinstaller puis taper la commande:
 pyinstaller --nocommand --onefile stats.py
Cette commande créera plusieurs dossier, l'éxecutable se trouvera au niveau du dossier dist.




---Utilisation du logiciel----

Ce logiciel a été développé afin de génerer des fichier Matbase, Matfinale et MatInter, pour les tâches VOL et Hoppy. Pour chaque fichier, trois sont génerés dépendant de la nature de la réponse (0, D ou M). Ce sont des fichiers Excel (.xlsx)
Mode d'emploi: 
-Il faut tout d'abord sélectionner l'emplacement de destination, où seront stockés les fichiers générés par le programme. Ceux-ci seront stockés dans un dossier nommé par défaut created_files. Il faut donc vérifier l'absence d'un dossier de ce nom dans l'emplacement sélectionné, sinon il sera écrasé et remplacé par le nouveau dossier généré.

-Ensuite, il faut mettre en entrée les fichiers CLAN représentant chacun un sujet. Pour avoir l'intégralité des informations dans le fichier, il faut veiller à insérer l'ensemble des fichiers CLAN. Le bouton à utiliser pour insérer les fichiers est le bouton "Explorer".

-Dans le cas d'une erreur, le bouton "Tout effacer" permet d'annuler l'importation en cours et de réinitialiser le répertoire de fichiers CLAN sélectionnés.
-Quand la sélection est faite, il suffit d'appuyer sur le bouton exécuter afin de commencer le programme.

Le logiciel est muni d'un correcteur de fichiers CLAN. Il traversera alors chaque fichier afin de vérifier si aucune des erreurs suivantes n'est constatée:
-Erreur sur le nombre de champs ou division erronée des champs
-Présence d'un caractère interdit lors du codage (",;!?.<>[]@")
-Présence d'un espace dans un des champs
-Absence du dernier caractère "{"
Si une erreur est constatée, le logiciel vous signalera le nom du fichier et une liste des erreurs présentes réferencées par lignes. Dans le cas où plusieurs fichiers sont erronées, des fenêtres contenant ces messages d'erreur s'ouvrent au fur et à mesure. Ainsi, pour passer aux erreurs présentes dans un autre fichier, il suffit de fermer la fenêtre contenant les erreurs du fichier précédent.
Sinon, le programme fait son éxécution et génère les fichiers Matbase.

Pour le choix de la tâche VOL, l'output sera le suivant:
-Un fichier dump qui divise chaque ligne %cod des fichiers CLAN en champs d'intérêt par sujet et par stimuli, qui seront réutilisés par la suite pour génerer les fichiers suivants.
-Des fichiers Matbase 0, D et M.
-Des fichiers Matinter 0, D et M classant et dénombrant les réponses selon leur place dans la phrase (verbe et périphérie)
-Des fichiers Matinterbis 0, D et M classant les réponses selon leur place dans la phrase (verbe et périphérie). Ici, on ne compte plus le nombre exact de réponses, mais le nombre de type de réponses.
-Des fichiers Matfinale 0,D et M classant et dénombrant les réponses quelle que soit leur place dans la phrase.
-Des fichiers Matfinalebis 0, D et M classant les réponses quelle que soit leur place dans la phrase (verbe et périphérie). Ici, on ne compte plus le nombre exact de réponses, mais le nombre de type de réponses.
