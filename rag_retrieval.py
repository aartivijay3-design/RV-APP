"""
rag_retrieval.py — BAWA Reference Library Retrieval

Loads reference_library.json, builds sentence-transformer embeddings on first
run (saved to reference_library.npy so it only runs once), then provides
retrieve() to find the most relevant German reference texts for a given query.

Usage:
    from rag_retrieval import retrieve
    results = retrieve("Kinkaku-ji Kyoto Japan", top_k=5)
    # returns list of {"type", "text", "destination", "location_heading", ...}
"""

import json
import numpy as np

from paths import DATA_DIR
import github_store

# ── Paths ─────────────────────────────────────────────────────────────────────
_JSON_PATH  = DATA_DIR / "reference_library.json"
_INDEX_PATH = DATA_DIR / "reference_library.npy"   # cached embeddings

# ── Lazy-loaded globals ────────────────────────────────────────────────────────
_model    = None   # SentenceTransformer model
_entries  = None   # list of dicts from JSON
_embeddings = None # np.ndarray, shape (N, dim)


def _load():
    """Load model + index (build index if not cached yet)."""
    global _model, _entries, _embeddings

    if _embeddings is not None:
        return  # already loaded

    print("[RAG] Loading sentence-transformer model …", flush=True)
    from sentence_transformers import SentenceTransformer
    _model = SentenceTransformer("paraphrase-multilingual-MiniLM-L12-v2")
    # ↑ 120 MB, multilingual (handles both German and English queries/texts),
    #   runs on CPU, no GPU needed, free.

    print("[RAG] Loading reference library …", flush=True)
    data = json.loads(_JSON_PATH.read_text(encoding="utf-8"))
    _entries = data["entries"]

    # On a host with no persistent disk, pull whatever was last built instead
    # of rebuilding from scratch every cold start — building takes ~2-4 min,
    # which combined with a free-tier wake-up delay would make the first
    # request after any idle period unacceptably slow. No-op if GitHub sync
    # isn't configured.
    if not _INDEX_PATH.exists():
        github_store.pull("data/reference_library.npy", _INDEX_PATH)

    if _INDEX_PATH.exists():
        print("[RAG] Loading cached embeddings …", flush=True)
        _embeddings = np.load(str(_INDEX_PATH))
        if len(_embeddings) != len(_entries):
            print("[RAG] Cache mismatch — rebuilding …", flush=True)
            _embeddings = None

    if _embeddings is None:
        print(f"[RAG] Building embeddings for {len(_entries)} entries "
              f"(one-time, ~2–4 min) …", flush=True)
        texts = [_entry_to_query_text(e) for e in _entries]
        _embeddings = _model.encode(
            texts,
            batch_size=64,
            show_progress_bar=True,
            normalize_embeddings=True,   # cosine similarity = dot product
        )
        np.save(str(_INDEX_PATH), _embeddings)
        print("[RAG] Embeddings saved to reference_library.npy", flush=True)
        github_store.push("data/reference_library.npy", _INDEX_PATH, "Update cached reference embeddings")

    print(f"[RAG] Ready — {len(_entries)} entries indexed.", flush=True)


def _entry_to_query_text(entry: dict) -> str:
    """
    Produce the string that gets embedded for each entry.
    Concatenating all searchable fields gives better recall.
    """
    parts = [
        entry.get("destination", ""),
        entry.get("location_heading", ""),
        entry.get("hotel_name", ""),
        entry.get("text", ""),
    ]
    return " | ".join(p for p in parts if p)


# Cosine similarity below which a "match" is treated as no match at all.
#
# Travel prose embeds into a narrow band — every entry in the corpus is
# topically "a German paragraph about sightseeing", so even completely
# unrelated pairs score high. Measured against a real Baltikum itinerary:
# "Besuch der Tori Cider Farm" (Estonia) matched a Riga farmers-market
# paragraph at 0.675, and "Rückflug" matched a hotel description at 0.688 —
# while the one genuinely correct match in the whole trip ("Mittagessen im
# Restaurant Lore Bistro" → the stored Lore Bistro line) scored 0.898. A
# threshold anywhere below ~0.85 therefore admits mostly noise.
#
# Deliberately biased toward returning nothing: a dropped style reference
# just means the model writes the paragraph from its own knowledge (which
# the prompt already handles), whereas a wrong one puts another trip's
# facts into a client-facing document — the failure this threshold exists
# to prevent (an unrelated Riga restaurant description appeared in a
# Tallinn day this way).
MIN_SEMANTIC_SCORE = 0.85


def retrieve(
    query: str,
    destination: str = "",
    top_k: int = 6,
    german_only: bool = True,
    min_text_len: int = 60,
    min_score: float = MIN_SEMANTIC_SCORE,
) -> list[dict]:
    """
    Find the most relevant reference entries for a query.

    Args:
        query:        Free-text query, e.g. "Kinkaku-ji golden pavilion Kyoto"
        destination:  If given, filters to entries with matching destination
                      (case-insensitive substring match). Pass "" to search all.
        top_k:        Number of results to return.
        german_only:  If True, only return entries with language == "de".
        min_text_len: Discard entries whose text is shorter than this (noise).
        min_score:    Minimum cosine similarity to count as a match at all
                      (see MIN_SEMANTIC_SCORE). Pass 0.0 to disable.

    Returns:
        List of entry dicts, ordered by relevance (most relevant first), each
        with its similarity in "score". Empty if nothing clears min_score.
    """
    _load()

    # ── Filter candidates ─────────────────────────────────────────────────────
    candidates = []
    candidate_indices = []
    dest_lower = destination.lower()

    for i, e in enumerate(_entries):
        if german_only and e.get("language") != "de":
            continue
        if len(e.get("text", "")) < min_text_len:
            continue
        if dest_lower and dest_lower not in e.get("destination", "").lower():
            continue
        candidates.append(e)
        candidate_indices.append(i)

    if not candidates:
        return []

    # ── Embed query ───────────────────────────────────────────────────────────
    q_vec = _model.encode(
        query,
        normalize_embeddings=True,
    )  # shape (dim,)

    # ── Score candidates ──────────────────────────────────────────────────────
    candidate_embeddings = _embeddings[candidate_indices]   # shape (M, dim)
    scores = candidate_embeddings @ q_vec                   # cosine similarity

    # ── Top-k ─────────────────────────────────────────────────────────────────
    # Always sort by score — the small-candidate-set path used to return
    # entries in corpus order, so the "most relevant first" contract silently
    # didn't hold for narrow destinations and callers taking results[0] got an
    # arbitrary entry rather than the best one.
    if len(scores) <= top_k:
        top_indices = sorted(range(len(scores)), key=lambda i: scores[i], reverse=True)
    else:
        top_indices = np.argpartition(scores, -top_k)[-top_k:]
        top_indices = sorted(top_indices, key=lambda i: scores[i], reverse=True)

    return [
        {**candidates[i], "score": float(scores[i])}
        for i in top_indices
        if scores[i] >= min_score
    ]


def retrieve_for_day(
    destination: str,
    activities: list[str],
    hotels: list[str] = None,
    top_k_sightseeing: int = 5,
    top_k_hotels: int = 2,
) -> str:
    """
    High-level helper: given a destination and list of activities/sights for
    one day, retrieves the most relevant German reference texts and formats
    them as a ready-to-inject prompt block.

    Args:
        destination:         e.g. "Japan"
        activities:          e.g. ["Kinkaku-ji", "Arashiyama Bamboo Grove",
                                   "Nishiki Market"]
        hotels:              e.g. ["Six Senses Kyoto"]  (optional)
        top_k_sightseeing:   how many sightseeing reference texts to retrieve
        top_k_hotels:        how many hotel reference texts to retrieve

    Returns:
        A formatted string block ready to append to the user prompt.
    """
    _load()
    blocks = []

    # ── Sightseeing ───────────────────────────────────────────────────────────
    if activities:
        query = f"{destination} {' '.join(activities)}"
        sight_results = retrieve(
            query=query,
            destination=destination,
            top_k=top_k_sightseeing,
            german_only=True,
        )
        if sight_results:
            blocks.append("=== BAWA REFERENCE TEXTS — SIGHTSEEING ===")
            blocks.append(
                "The following are real German texts from confirmed BAWA itineraries "
                "for similar destinations and sights. Reuse phrasing, sentence rhythm, "
                "and vocabulary where it fits. Adapt content to match the actual "
                "activities in this DMC — do not copy facts that don't apply."
            )
            for r in sight_results:
                label = f"[{r['destination']} / {r.get('location_heading','')}]"
                blocks.append(f"{label}\n{r['text']}")
            blocks.append("")

    # ── Hotels ────────────────────────────────────────────────────────────────
    if hotels:
        for hotel_name in hotels:
            query = f"{hotel_name} {destination} hotel"
            hotel_results = retrieve(
                query=query,
                destination=destination,
                top_k=top_k_hotels,
                german_only=True,
            )
            # Prefer exact hotel name matches
            exact = [r for r in hotel_results
                     if hotel_name.lower() in r.get("hotel_name", "").lower()
                     or hotel_name.lower() in r.get("text", "").lower()]
            results_to_use = exact if exact else hotel_results[:1]

            if results_to_use:
                if not any("HOTEL" in b for b in blocks):
                    blocks.append("=== BAWA REFERENCE TEXTS — HOTELS ===")
                    blocks.append(
                        "The following are real BAWA hotel descriptions. "
                        "If the hotel matches, reuse the description directly. "
                        "If it's a different hotel, use it as a style reference only."
                    )
                for r in results_to_use:
                    label = (f"[{r.get('hotel_name', r.get('location_heading',''))} "
                             f"/ {r['destination']}]")
                    blocks.append(f"{label}\n{r['text']}")
                blocks.append("")

    if not blocks:
        return ""

    return "\n".join(blocks)


# ── CLI test ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("\n--- Testing retrieve() ---\n")
    results = retrieve(
        query="Kinkaku-ji goldener Pavillon Kyoto",
        destination="Japan",
        top_k=3,
    )
    for r in results:
        print(f"[{r['destination']} / {r.get('location_heading','')}]")
        print(r['text'])
        print()

    print("\n--- Testing retrieve_for_day() ---\n")
    block = retrieve_for_day(
        destination="Japan",
        activities=["Arashiyama Bamboo Grove", "Kinkaku-ji"],
        hotels=["Six Senses Kyoto"],
    )
    print(block)
