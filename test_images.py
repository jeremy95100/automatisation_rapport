"""
Script de test pour vérifier l'intégration des images dans les documents DOCX.
"""
from pathlib import Path
from remplace_rapport import process_document

def test_image_integration():
    """
    Test l'intégration d'images dans un document DOCX.

    Prérequis:
    - Un fichier test.docx avec des placeholders et des titres
    - Une image de test (test_image.jpg ou test_image.png) dans le dossier
    """

    # Chemins
    input_doc = Path("test.docx")
    output_doc = Path("test_avec_images.docx")
    test_image = None

    # Chercher une image de test
    for ext in ['.jpg', '.jpeg', '.png', '.gif']:
        img_path = Path(f"test_image{ext}")
        if img_path.exists():
            test_image = str(img_path)
            break

    if not input_doc.exists():
        print("❌ Fichier test.docx introuvable")
        return False

    if not test_image:
        print("⚠️  Aucune image de test trouvée (test_image.jpg, .png, etc.)")
        print("   Le test va continuer sans images")
    else:
        print(f"✓ Image de test trouvée: {test_image}")

    # Configuration du test
    mapping = {
        "{nom}": "Test User",
        "{date}": "2024-01-15",
        "{daterap}": "2024-01-20",
        "{detaille}": "détaillées du système"
    }

    # Configuration des images (à adapter selon vos titres réels)
    images_config = {}
    if test_image:
        # Exemple 1: Insertion après un titre spécifique
        images_after_headings = {
            "Résultats": test_image,  # Remplacez par un vrai titre de votre document
        }

        # Exemple 2: Insertion via marqueur [[IMG:graphique]]
        images_at_markers = {
            "graphique": test_image,
        }

        # Exemple 3: Insertion après un paragraphe spécifique
        images_after_paragraphs = {
            "Voir le graphique ci-dessous:": test_image,
        }

        # Configuration des tailles d'images
        images_sizes = {
            "Résultats": 4.0,  # 4 pouces de large
            "graphique": 3.5,
        }

        # Configuration du texte avant/après les images
        image_texts = {
            "Résultats": {
                "before": "Figure 1:",
                "after": "",
                "position": "before"
            },
            "graphique": {
                "before": "",
                "after": "(Source: Analyse interne)",
                "position": "after"
            }
        }
    else:
        images_after_headings = {}
        images_at_markers = {}
        images_after_paragraphs = {}
        images_sizes = {}
        image_texts = {}

    print("\n📄 Traitement du document...")
    print(f"   Input:  {input_doc}")
    print(f"   Output: {output_doc}")

    try:
        process_document(
            input_path=str(input_doc),
            output_path=str(output_doc),
            mapping_override=mapping,
            decisions_override=None,  # Utilise les phrases par défaut
            interactive=False,
            images_after_headings=images_after_headings,
            images_after_paragraphs=images_after_paragraphs,
            images_at_markers=images_at_markers,
            image_width_inches=3.0,
            images_after_headings_sizes=images_sizes,
            images_at_markers_sizes=images_sizes,
            image_texts=image_texts
        )

        if output_doc.exists():
            print(f"\n✓ Document généré avec succès: {output_doc}")
            print(f"  Taille: {output_doc.stat().st_size / 1024:.1f} KB")
            return True
        else:
            print("\n❌ Le document de sortie n'a pas été créé")
            return False

    except Exception as e:
        print(f"\n❌ Erreur lors du traitement: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("Test d'intégration des images dans les documents DOCX")
    print("=" * 60)

    success = test_image_integration()

    print("\n" + "=" * 60)
    if success:
        print("✓ Test réussi!")
        print("\nPour utiliser les images dans votre document:")
        print("1. Via API: utilisez les champs images_after_headings,")
        print("   images_after_paragraphs, ou images_at_markers")
        print("2. Via marqueurs: ajoutez [[IMG:cle]] dans votre document")
        print("   Word, puis mappez 'cle' vers le chemin de l'image")
    else:
        print("❌ Test échoué - voir les erreurs ci-dessus")
    print("=" * 60)
