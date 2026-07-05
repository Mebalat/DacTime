# DacTime

Application Windows de redaction et d'impression de courriers postaux.

Developpee en Python/Tkinter par Clement LATTAR.

## Fonctionnalites

- Redaction de courriers sur plusieurs onglets
- Champs expediteur et destinataire
- Export PDF et impression directe
- Impression d'enveloppes DL (Ctrl+E)
- Correcteur orthographique (LanguageTool, F7)
- Themes clair et sombre
- Zoom texte, rechercher/remplacer
- Sauvegarde au format .dactime
- Historique des destinataires

## Raccourcis clavier

| Raccourci | Action |
|-----------|--------|
| Ctrl+N | Nouveau courrier |
| Ctrl+S | Enregistrer |
| Ctrl+P | Imprimer |
| Ctrl+E | Imprimer enveloppe DL |
| Ctrl+F | Rechercher |
| F1 | Raccourcis clavier |
| F7 | Correcteur orthographique |

## Installation

Telecharger `DacTime_Setup_1.2.exe` dans les [Releases](../../releases) et lancer l'installateur.

## Compiler depuis les sources

Prerequis : Python 3.12, PyInstaller, NSIS

```
pip install pyinstaller pillow reportlab pywin32
pyinstaller DacTime.spec -y
makensis DacTime_setup.nsi
```

## Licence

MIT — voir [LICENSE](LICENSE)
