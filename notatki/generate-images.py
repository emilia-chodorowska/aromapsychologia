#!/usr/bin/env python3
"""
Generowanie grafik do prezentacji Trening Węchowy — Wellness
Używa Google Gemini API (Nano Banana = gemini-2.5-flash-image)
"""

import os
import sys
from google import genai
from google.genai import types

# --- Konfiguracja ---
API_KEY = os.environ.get("GEMINI_API_KEY")
if not API_KEY:
    # Spróbuj wczytać z .env w katalogu projektu
    env_path = os.path.join(os.path.dirname(__file__), "..", ".env")
    if os.path.exists(env_path):
        with open(env_path) as f:
            for line in f:
                if line.startswith("GEMINI_API_KEY="):
                    API_KEY = line.strip().split("=", 1)[1]
if not API_KEY:
    print("Brak GEMINI_API_KEY! Ustaw zmienną lub dodaj do .env")
    sys.exit(1)

MODEL = "gemini-2.5-flash-image"
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

# --- Styl wspólny dla wszystkich promptów ---
STYLE = (
    "Minimal flat illustration, soft sage green (#a7c5a9) and slate grey (#475569) accents, "
    "white background, clean wellness aesthetic, no text, no watermarks, high quality"
)

# --- Lista grafik ---
IMAGES = [
    {
        "filename": "grafika-grypa.png",
        "prompt": (
            "Photorealistic portrait of a person with a cold, holding their nose, visibly congested "
            "and uncomfortable, red irritated nose area, tissues nearby, soft indoor lighting, "
            "shallow depth of field, no text"
        ),
    },
    {
        "filename": "grafika-covid.png",
        "prompt": (
            "Photorealistic 3D render of SARS-CoV-2 coronavirus particles, detailed spike proteins, "
            "red-orange virus spheres floating against a soft blurred dark background, scientific "
            "visualization style, dramatic lighting, shallow depth of field, no text"
        ),
    },
    {
        "filename": "grafika-technika-prawidlowa.png",
        "prompt": (
            f"A person gently sniffing a small dark glass jar held close to their nose, "
            f"eyes closed, taking short delicate breaths. Calm, meditative posture. {STYLE}"
        ),
    },
    {
        "filename": "grafika-technika-blad.png",
        "prompt": (
            f"A person taking one deep, long breath through the nose with a jar far from face, "
            f"head tilted back. A subtle red X mark indicating wrong technique. {STYLE}"
        ),
    },
    {
        "filename": "grafika-dzienniczek.png",
        "prompt": (
            f"An open notebook or journal with handwritten notes and a pen lying beside it, "
            f"small essential oil bottles nearby on a clean desk. {STYLE}"
        ),
    },
    {
        "filename": "grafika-neuroplastycznosc.png",
        "prompt": (
            f"A human brain with glowing neural pathways reconnecting and forming new connections. "
            f"Soft bioluminescent glow on the neural network. {STYLE}"
        ),
    },
    {
        "filename": "grafika-wachanie.png",
        "prompt": (
            "Photorealistic portrait of a woman in cozy casual clothes sitting on a sofa at home, "
            "eyes closed, gently smelling a small dark amber glass jar held close to her nose, "
            "warm soft natural light, cozy living room background with plants, focused concentrated expression, "
            "shallow depth of field, no text"
        ),
    },
    {
        "filename": "grafika-zasady.png",
        "prompt": (
            "Fotorealistyczne ujęcie z góry na jasnym drewnianym stole: małe ciemne szklane słoiczki "
            "z korkowymi zatyczkami, świeża cytryna, płatki róży, gałązki eukaliptusa, laski cynamonu, "
            "otwarty notes z długopisem, miękkie naturalne światło dzienne, ciepła estetyka wellness, "
            "mała głębia ostrości, bez tekstu"
        ),
    },
]


def main():
    client = genai.Client(api_key=API_KEY)

    total = len(IMAGES)
    for i, img in enumerate(IMAGES, 1):
        filepath = os.path.join(OUTPUT_DIR, img["filename"])

        # Pomiń jeśli plik już istnieje
        if os.path.exists(filepath):
            print(f"  [{i}/{total}] {img['filename']} — już istnieje, pomijam")
            continue

        print(f"  [{i}/{total}] Generuję {img['filename']}...")

        try:
            response = client.models.generate_content(
                model=MODEL,
                contents=[img["prompt"]],
                config=types.GenerateContentConfig(
                    response_modalities=["IMAGE", "TEXT"],
                ),
            )

            saved = False
            for part in response.candidates[0].content.parts:
                if part.inline_data is not None:
                    image = part.as_image()
                    image.save(filepath)
                    print(f"           ✓ Zapisano: {filepath}")
                    saved = True
                    break

            if not saved:
                print(f"           ✗ Brak obrazu w odpowiedzi!")
                if response.candidates[0].content.parts:
                    for part in response.candidates[0].content.parts:
                        if part.text:
                            print(f"           Odpowiedź: {part.text[:200]}")

        except Exception as e:
            print(f"           ✗ Błąd: {e}")

    print(f"\nGotowe! Sprawdź folder: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
