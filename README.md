# DacTime

Outil personnel pour rédiger et imprimer des courriers postaux sous Windows.

Développé parce que Word est trop lourd et le Bloc-notes pas assez. Peut servir à d'autres dans le même cas.

## Ce que ça fait

- Rédaction sur plusieurs onglets (un courrier par onglet)
- Champs expéditeur et destinataire, historique des destinataires
- Impression directe ou export PDF
- Impression d'enveloppes DL (Ctrl+E) via PDF
- Correcteur orthographique intégré via LanguageTool (F7)
- Thèmes clair et sombre, zoom texte
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

Télécharger `DacTime_Setup_1.2.exe` dans les [Releases](../../releases) et lancer l'installateur.

Windows peut afficher un avertissement à l'ouverture : cliquer **Informations complémentaires** puis **Exécuter quand même**.

## Compiler depuis les sources

Prérequis : Python 3.12, PyInstaller, NSIS

```
pip install pyinstaller pillow reportlab pywin32
pyinstaller DacTime.spec -y
makensis DacTime_setup.nsi
```

## Licence

MIT, voir [LICENSE](LICENSE).

---

Si ce projet te sert, un café est le bienvenu : bouton Sponsor en haut à droite.
