# Gain de temps — avant / après

L'impact business d'une automatisation se mesure en **temps rendu aux équipes**.

## Le rapport, à la main (avant)

Un reporting de ventes mensuel refait manuellement, c'est typiquement :

| Étape manuelle | Temps |
|---|---|
| Extraire les données (export, copier-coller) | ~10 min |
| Construire les tableaux croisés / agrégations | ~15 min |
| Mise en forme (couleurs, formats €, bordures) | ~10 min |
| Créer les graphiques | ~8 min |
| Vérifications + envoi | ~5 min |
| **Total par rapport** | **~48 min** |

Fait **chaque semaine** → ~3,2 h/mois → **~38 heures par an**. Et le risque d'erreur
humaine à chaque copier-coller.

## Le rapport, automatisé (après)

```bash
python src/generate_report.py     # < 2 secondes
```

| Indicateur | Avant | Après |
|---|---|---|
| Temps par rapport | ~48 min | **< 1 min** (0 manuel) |
| Temps par an | ~38 h | **~0 h** |
| Risque d'erreur | élevé (manuel) | **nul** (déterministe) |
| Déclenchement | humain | **planifié** (cron / Task Scheduler) |

➡️ **~38 heures/an économisées** + un rapport toujours identique, fiable et à l'heure.

## Planification

**Windows (Planificateur de tâches)** — action :
```
python  C:\...\projet-06-automatisation-reporting-excel\src\generate_report.py
```
déclenchée chaque lundi à 8h, par exemple.

**Linux/macOS (cron)** :
```
0 8 * * 1  cd /chemin/projet-06 && python src/generate_report.py
```
