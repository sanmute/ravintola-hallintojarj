"""
translator.py — Paikallinen (offline) englanti->suomi-kääntäjä.

Käyttää suoraan ctranslate2- ja sentencepiece-kirjastoja Argos Translaten
en->fi-mallilla (models/translate-en_fi/), EI argostranslate-pakettia itseään.
Syy: argostranslate lataa oletuksena raskaat lauseenrajaus-riippuvuudet
(torch/spacy/stanza, yhteensä ~1.2 GB) pelkkää pitkien kappaleiden
lauseenjakoa varten. Tässä sovelluksessa käännetään vain lyhyitä kenttiä
(reseptin nimi, yksittäinen raaka-aine, muutama ohjelause) — rivinvaihdolla
jaettu käännös per rivi riittää, joten koko riippuvuuspino ei ole tarpeen.
Jäljelle jäävä jalanjälki: ctranslate2 (~60 MB) + sentencepiece (~5 MB) +
malli (~80 MB) sen sijaan että ~1.2 GB.
"""

import os
import threading

_MODEL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'models', 'translate-en_fi')

_translator = None
_sp = None
_lock = threading.Lock()
_load_error = None


def _ensure_loaded():
    global _translator, _sp, _load_error
    if _translator is not None or _load_error is not None:
        return
    with _lock:
        if _translator is not None or _load_error is not None:
            return
        try:
            import ctranslate2
            import sentencepiece as spm
            _sp = spm.SentencePieceProcessor(
                model_file=os.path.join(_MODEL_DIR, 'sentencepiece.model'))
            _translator = ctranslate2.Translator(
                os.path.join(_MODEL_DIR, 'model'), device='cpu')
        except Exception as e:
            _load_error = str(e)


def translate_en_to_fi(text):
    """Käännä englanninkielinen teksti suomeksi. Käsittelee monirivisen
    tekstin rivi kerrallaan (ei lauseenrajausta — riittää lyhyille
    resepti-/raaka-ainekentille). Palauttaa (translated_text, error)."""
    if not text or not text.strip():
        return '', None

    _ensure_loaded()
    if _load_error is not None:
        return None, f'Kääntäjää ei saatu ladattua: {_load_error}'

    lines = text.split('\n')
    translated_lines = []
    for line in lines:
        if not line.strip():
            translated_lines.append('')
            continue
        tokens = _sp.encode(line, out_type=str)
        result = _translator.translate_batch(
            [tokens], beam_size=4, num_hypotheses=1, replace_unknowns=True)
        out_tokens = result[0].hypotheses[0]
        decoded = _sp.decode_pieces(out_tokens).replace('▁', ' ').replace('_', ' ')
        if decoded.startswith(' '):
            decoded = decoded[1:]
        translated_lines.append(decoded)

    return '\n'.join(translated_lines), None
