# ConAn - Contrôle Alma

[![Abandonned](https://img.shields.io/badge/Maintenance%20Level-Abandoned-orange.svg)](https://gist.github.com/cheerfulstoic/d107229326a01ff0f333a1d3476e068d)

ConAn est un outil visant originellement à contrôler des données provenant d'Alma. En pratique, il permet également de donner certaines statistiques ou informations sur le fonds renseigné.

**Évitez d'avoir d'autres fichiers Excel ouverts pendant l'analyse (dans le cas où une erreur de programmation pourrait faire intéragir ConAn avec des fichiers non prévus).**

## Initialisation

Exportez d'Alma une liste de Titres physiques, renommez-la `export_Alma_ConAn.xlsx` et placez-la dans le même dossier que `ConAn.xlsm` (ConAn dans le reste de la documentation).

Allumez ConAn, choisissez la feuille `Introduction` et remettez à zéro les données. Sélectionnez la bibliothèque en `H4` (en appuyant sur Alt + flèche du bas, une liste déroulante s'affichera) puis le contrôle à effectuer en `H2` (même manipulation pour voir la liste).

Lancez ensuite l'analyse.

### Ajouter une bibliothèque

La liste des bibliothèques se trouvent en colonne `P` (masquée par défaut). La liste déroulante se mettra automatiquement à jour après un ajout (jusqu'à 99 bibliothèques. Au-delà, il faudra modifier la formule).

## Les analyses

### CA1 : analyse cote / holding

Pas encore implémentée, elle servira à comparer si la cote d'un document correspond à la holding associée, utilisant un fonctionnement similaire (mais plus simple) que [Louise](https://github.com/Alban-Peyrat/Louise). Le développement est pour l'instant uniquement prévu pour la BUSVS.

### CA2 : statistiques d'âge en prenant en compte les exemplaires

Calcule l'âge moyen et l'âge médian d'un fonds en fonction du nombre d'exemplaires (entre 1900 et 2030). C'est une version améliorée de [l'analyse CS4 de ConStance](https://github.com/Alban-Peyrat/ConStance#cs4-non-pref--statistiques-d%C3%A2ge-champs-210-214) qui ne prend pas compte les exemplaires.

Pour chaque titre dans la liste, ConAn prélève le PPN, l'année la plus récente présente dans la colonne `Créateur / Éditeur` et le nombre d'exemplaires.

Pour isoler l'année, ConAn :
* supprime tout ce qui se trouve avant la première parenthèse (supposément la partie consacrée au créateur) ;
* divise en plusieurs parties ce qu'il reste, en utilisant le point-virgule comme séparateur (les différentes 210/214 semblent être séparées par des points-virgules dans l'export Alma) ;
* pour chacune de ces parties, il cherche la position de la dernière virgule. Si celle-ci existe et se trouve avant les derniers 4 caractères (empêche des erreurs de catalogage de nuire à l'analyse), il conserve tout ce qui se trouve derrière la virgule, sinon il conserve l'intégralité de la partie ;
* supprime alors les espaces, `DL`, `C`, `cop`, `P`, parenthèses ainsi que les caractères suivants : `. , - ?` ; 
* puis regarde si ce qu'il reste peut être considérer comme un nombre ;
* si ce n'est pas le cas, ConAn passe un par un les caractères restants en ne conservant que les chiffres. Au fur et à mesure de ce processus, à chaque fois qu'il a en mémoire 8 caractères, il les divise en deux paquets de 4 chiffres et compare lequel des deux est le plus grande sous forme de nombre (ce qui vise à éviter les erreurs de dépassement de capacité et à ne conserver que l'année la plus grande). Si au terme de l'analyse manuelle, aucun nombre ne reste, il attribue alors 0 à la valeur de l'année ;
* ConAn regarde ensuite si l'année qu'il a conservée se trouve entre 1900 et 2030 (exclus). Si ce n'est pas le cas, il regarde si les quatre derniers chiffres sont situés dans cet intervalle. Si ce n'est toujours pas le cas, il regarde si les quatre premiers chiffres sont situés dans l'intervalle. Si aucun cas ne correspond, il attribue alors la valeur 0 à l'année. Cette partie permet d'exclure des nombres qui seraient trop extrêmes et donc probablement des erreurs dans l'isolation de la cote ;
* enfin, il vérifie s'il a déjà une année en mémoire pour ce PPN, si ce n'est pas le cas, il conserve celle-ci, sinon, il compare les deux et conserve la plus récente ;
* effectue ensuite la même analyse sur toutes les autres parties ;
* si à la fin il n'a pas d'année ou que celle-ci équivaut à 0, il ne renvoie aucune année dans la page de résultat.

Pour calculer le nombre d'exemplaires, ConAn :
* divise la colonne de disponibilité en plusieurs parties, chacune correspondant à une holding ;
* localise ensuite, pour chaque holding correspondant à la bibliothèque sélectionnée et qui n'est pas une holding associée, le terme `exemplaire` (s'il n'est pas présent, il passe à la prochaine holding, ce implique que les exemplaires ayant une mention de tomaison dans l'exemplaire sont exclus) ;
* localise alors la première parenthèse et ne conserve que ce qui se trouve entre celle-ci et le terme `exemplaire`. Il répète cette opération tant que l'écart entre ces deux points est supérieur à 7. En effet lorsque que l'écart est inférieur à 7, nous avons forcément l'information sur les exemplaires dans ce qui reste (fonctionne jusqu'à 999 exemplaires) ;
* supprime ensuite espaces et lettres `e`, ce qui laisse supposément uniquement un nombre, mais, dans le cas où le reste ne serait pas un nombre, il passe un par un les caractères pour conserver uniquement les chiffres ;
* additionne ensuite ce nombre au total d'exemplaires pour ce PPN (0 par défaut) et passe à la holding suivante.

Une fois tous les PPN traités, il calcule la moyenne et la médiane en prenant en compte le nombre d'exemplaires. Concernant la page de résultat, les PPN sont triés par année, ceux n'ayant pas d'année associée sont triés au fonds et colorés de rouge. Il comptabilise également le nombre de titres exclus et le nombre d'exemplaires exclus.

### CA3 : détection de multiples éditions d'un même titre

_Pas entièrement à jour, notamment prend en compte les 979 en début d'ISBN 13_

_A pour alternative [CS6 de ConStance](https://github.com/Alban-Peyrat/ConStance#cs6--d%C3%A9tection-de-multiples-%C3%A9ditions-dun-m%C3%AAme-titre). Chacun des deux à ses avantages et ses inconvénients, la détection initiale via clef de titre est différente._

Détermine les titres qui ont possiblement deux éditions dans la même liste.

Pour chaque titre dans la liste, ConAn prélève le PPN et les colonnes du titre, de l'édition, de l'ISBN 13 et de l'ISBN 10.

Il va ensuite générer une clef pour chaque titre :
* il ne conserve de la donnée originelle que la partie du titre se trouvant avant ` / ` (espaces avant et après), soit très généralement le titre propre (sauf si la dite chaîne de caractère se trouve dans le titre propre) ;
* il supprime ensuite `[Texte imprimé]`, les articles définis et indéfinis, les parenthèses, crochets et la liste de caractères suivants : `, : ; . '` et remplace les tirets par des espaces ;
* il divise alors son nouveau titre en autant de mots qu'il détecte (en utilisant les espaces comme séparateurs) ;
* pour le premier mot, il conserve les 4 premiers caractères, puis les deux premiers caractères des trois mots suivants (s'ils existent), séparant les chaînes de caractères par des _underscores_ (`_`) ;
* la clef est ensuite passée en majuscule (elle conserve les accents).

Ensuite, pour chaque PPN dont la colonne d'édition n'est pas vide, il compare pour l'intégralité des autres entrées de la liste si la clef est parfaitement égale à la sienne. Si c'est le cas, il conserve dans un tableau le PPN, l'ISBN 13 et l'ISBN 10 de cette entrée.

Une fois l'intégralité de la liste parcourue pour ce PPN, si ce tableau n'est pas vide, il génère une clef d'ISBN pour ce PPN :
* si l'ISBN 13 n'est pas vide et qu'il est sous sa forme avec des tirets, il en conserve les trois premières parties (c'est-à-dire le pays producteurs ou distributeur (_bookland_ pour les livres, soit 978 ou 979), le domaine ISBN et le numéro d'éditeur) ;
* s'il n'a pas pu récupérer via l'ISBN 13 la clef et que l'ISBN 10 n'est pas vide et qu'il est sous sa forme avec des tirets, il ajout `978-` (première partie de l'ISBN 13) aux deux premières parties de l'ISBN 10 ;
* sinon, la clef est une chaîne de caractères vide.

Une fois sa propre clef générée, pour chacune des entrées de son tableau (=des PPN potentiellement doublons d'édition), il génère la clef ISBN de celles-ci et la compare à celle du PPN originale, trois cas sont alors possibles :
* `corr. ISBN` : les clefs ISBN correspondent, la ligne du PPN original se colore en rouge ;
* `imp. ISBN` : une des deux clefs ISBN au moins est vide, ConAn n'a donc pas pu les comparer, la ligne se colore en orange/jaune ;
* `NO corr. ISBN` : les clefs ISBN ne correspondent pas, la ligne se colore en bleu.

ConAn indique enfin dans une dernière colonne le résultat de son analyse :
* soit `Aucune détection automatique`, colorée de vert ;
* soit `Double éd. possible` suivi de tous les PPN détectés avec, pour chacun d'entre eux, le résultat de la correspondance ISBN (la couleur de la ligne suit l'ordre de priorité susmentionné).

Note : c'est une détection automatique qui présente des limites, la liste des résultats peut ne pas détecter certains cas tout comme elle peut détecter des faux positifs (la comparaison ISBN cherche à diminuer ces faux-positifs).
