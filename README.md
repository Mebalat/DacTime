# DacTime

Outil personnel pour rediger et imprimer des courriers postaux sous Windows.

Developpe parce que Word est trop lourd et le Bloc-notes pas assez. Peut servir a d'autres dans le meme cas.

## Ce que ca fait

- Redaction sur plusieurs onglets (un courrier par onglet)
- Champs expediteur et destinataire, historique des destinataires
- Impression directe ou export PDF
- Impression d'enveloppes DL (Ctrl+E) via PDF
- Correcteur orthographique integre via LanguageTool (F7)
- Themes clair et sombre, zoom texte
- Sauvegarde au format .dactime

## Raccourcis

| Raccourci | Action |
|-----------|--------|
| Ctrl+N | Nouveau courrier |
| Ctrl+S | Enregistrer |
| Ctrl+P | Imprimer |
| Ctrl+E | Enveloppe DL |
| Ctrl+F | Rechercher |
| F1 | Liste des raccourcis |
| F7 | Correcteur orthographique |

## Installation

Telecharger `DacTime_Setup_1.2.exe` dans les [Releases](../../releases) et lancer l'installateur.

Windows peut afficher un avertissement a l'ouverture : cliquer **Informations complementaires** puis **Executer quand meme**.

## Compiler depuis les sources

Prerequis : Python 3.12, PyInstaller, NSIS

```
pip install pyinstaller pillow reportlab pywin32
pyinstaller DacTime.spec -y
makensis DacTime_setup.nsi
```

## Licence

MIT — voir [LICENSE](LICENSE)

---

Si ce projet te sert, un cafe est le bienvenu — bouton Sponsor en haut a droite.
