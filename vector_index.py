"""
Semantic search index for OneNote pages using sentence-transformers.

Builds page-level embeddings and supports cosine similarity search.
Index is persisted to ~/.cache/onenote-mcp/ and incrementally updated
based on file modification times.
"""

import json
import logging
import os
from pathlib import Path

import numpy as np

log = logging.getLogger("onenote-mcp")

MODEL_NAME = "paraphrase-multilingual-MiniLM-L12-v2"
CACHE_DIR = Path.home() / ".cache" / "onenote-mcp"


class EmbeddingIndex:
    def __init__(self):
        self._model = None
        self._embeddings: np.ndarray | None = None  # shape (N, dim)
        self._metadata: list[dict] = []  # parallel to embeddings rows

    def _get_model(self):
        if self._model is None:
            from sentence_transformers import SentenceTransformer
            self._model = SentenceTransformer(MODEL_NAME)
        return self._model

    def build(self, notebooks: dict, parse_pages_fn) -> int:
        """
        Build or incrementally update the index from discovered notebooks.

        Args:
            notebooks: Output of _discover_notebooks().
            parse_pages_fn: Reference to _parse_pages function.

        Returns:
            Number of pages indexed.
        """
        self.load()

        # Build a map of file_path -> mtime from existing metadata
        existing = {}
        for i, m in enumerate(self._metadata):
            existing[m["file_path"]] = {
                "mtime": m["file_mtime"],
                "index": i,
            }

        # Collect all pages that need (re-)embedding
        new_chunks: list[dict] = []
        keep_indices: list[int] = []
        current_paths: set[str] = set()

        for nb_name, nb_info in notebooks.items():
            for sec_name, sec_info in nb_info["sections"].items():
                filepath = sec_info["latest"]
                fpath_str = str(filepath)
                current_paths.add(fpath_str)
                file_mtime = filepath.stat().st_mtime

                if fpath_str in existing and existing[fpath_str]["mtime"] == file_mtime:
                    # File unchanged — keep existing entries for this file
                    for i, m in enumerate(self._metadata):
                        if m["file_path"] == fpath_str:
                            keep_indices.append(i)
                    continue

                # File is new or changed — parse and prepare for embedding
                pages = parse_pages_fn(filepath)
                for page in pages:
                    full_text = "\n".join(page["texts"])
                    if not full_text.strip():
                        continue
                    chunk_text = f"{page['title']}\n{full_text}"
                    new_chunks.append({
                        "notebook": nb_name,
                        "section": sec_name,
                        "page_title": page["title"],
                        "text": full_text,
                        "file_path": fpath_str,
                        "file_mtime": file_mtime,
                        "_embed_text": chunk_text,
                    })

        # Remove entries for files that no longer exist
        for i, m in enumerate(self._metadata):
            if m["file_path"] not in current_paths and i in keep_indices:
                keep_indices.remove(i)

        # Build new embeddings for changed/new chunks
        if new_chunks:
            model = self._get_model()
            texts_to_embed = [c["_embed_text"] for c in new_chunks]
            log.info("Embedding %d new/changed pages...", len(texts_to_embed))
            new_embeds = model.encode(texts_to_embed, normalize_embeddings=True,
                                      show_progress_bar=False)
            new_embeds = np.array(new_embeds, dtype=np.float32)

            # Strip _embed_text from metadata
            for c in new_chunks:
                del c["_embed_text"]
        else:
            new_embeds = None

        # Merge: kept old entries + new entries
        if keep_indices and self._embeddings is not None:
            kept_embeds = self._embeddings[keep_indices]
            kept_meta = [self._metadata[i] for i in keep_indices]
        else:
            kept_embeds = np.empty((0, 384), dtype=np.float32)
            kept_meta = []

        if new_embeds is not None and len(new_embeds) > 0:
            self._embeddings = np.vstack([kept_embeds, new_embeds]) if len(kept_embeds) > 0 else new_embeds
            self._metadata = kept_meta + new_chunks
        else:
            self._embeddings = kept_embeds if len(kept_embeds) > 0 else None
            self._metadata = kept_meta

        total = len(self._metadata)
        log.info("Index contains %d pages total (%d new/changed, %d kept)",
                 total, len(new_chunks) if new_chunks else 0, len(keep_indices))

        if total > 0:
            self.save()

        return total

    def search(self, query: str, top_k: int = 20) -> list[dict]:
        """
        Search the index for pages semantically similar to query.

        Returns list of dicts with keys: notebook, section, page_title, text, score.
        """
        if self._embeddings is None or len(self._metadata) == 0:
            return []

        model = self._get_model()
        query_embed = model.encode([query], normalize_embeddings=True)
        query_embed = np.array(query_embed, dtype=np.float32)

        # Cosine similarity (embeddings are already normalized)
        scores = (self._embeddings @ query_embed.T).flatten()

        top_indices = np.argsort(scores)[::-1][:top_k]

        results = []
        for idx in top_indices:
            score = float(scores[idx])
            if score < 0.1:
                break
            m = self._metadata[idx]
            results.append({
                "notebook": m["notebook"],
                "section": m["section"],
                "page_title": m["page_title"],
                "text": m["text"],
                "score": score,
            })

        return results

    def save(self):
        """Persist index to disk."""
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        if self._embeddings is not None:
            np.save(CACHE_DIR / "embeddings.npy", self._embeddings)
        meta_path = CACHE_DIR / "metadata.json"
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(self._metadata, f, ensure_ascii=False)
        log.info("Saved index to %s (%d entries)", CACHE_DIR, len(self._metadata))

    def load(self):
        """Load index from disk if available."""
        embed_path = CACHE_DIR / "embeddings.npy"
        meta_path = CACHE_DIR / "metadata.json"
        if embed_path.exists() and meta_path.exists():
            try:
                self._embeddings = np.load(embed_path)
                with open(meta_path, "r", encoding="utf-8") as f:
                    self._metadata = json.load(f)
                log.info("Loaded index from %s (%d entries)", CACHE_DIR, len(self._metadata))
            except Exception as e:
                log.warning("Failed to load index: %s", e)
                self._embeddings = None
                self._metadata = []
        else:
            self._embeddings = None
            self._metadata = []
