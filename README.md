# Dakar Climate Intelligence Hub

Depot de travail pour le classeur **DAKAR_CLIMATE_INTELLIGENCE_HUB**.

Le projet contient des scripts de generation et d'analyse pour construire un hub Excel de suivi climatique pour Dakar : indicateurs, inventaire GES, qualite de l'air, mobilite, energie, dechets, financements, reporting et tableaux de bord.

## Contenu

- Scripts JavaScript `.mjs` pour generer les classeurs Excel et les apercus.
- Scripts Python `.py` pour analyser les sources documentaires et extraire des indices.
- Dossier `outputs/` avec les classeurs Excel generes, les apercus PNG et les index JSON.
- Fichiers de configuration Git et dependances Python minimales.

## Fichiers exclus

Le dossier `outputs/DOCUMENTS/` n'est pas versionne. Il contient des documents sources PDF/DOCX/PPTX potentiellement volumineux ou prives.

Le dossier `node_modules/`, les DLL locales et les fichiers temporaires Office sont egalement exclus.

## Execution

Les scripts JavaScript utilisent `@oai/artifact-tool`, fourni par l'environnement de travail Codex utilise pour generer ces classeurs.

Exemples :

```powershell
node build_base_par_annee.mjs
node build_dakar_climate_master_v6_simplifie.mjs
node render_apercus.mjs
```

Pour les scripts Python :

```powershell
pip install -r requirements.txt
python extract_document_evidence.py
```
