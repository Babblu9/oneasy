import fs from 'fs';

const MAX_CACHE_ITEMS = 3;
const cache = new Map();

export function workbookFingerprint(filePath) {
    const st = fs.statSync(filePath);
    return `${filePath}:${st.size}:${Math.trunc(st.mtimeMs)}`;
}

export function getCachedEngine(fingerprint) {
    if (!fingerprint) return null;
    const entry = cache.get(fingerprint);
    if (!entry) return null;

    // Touch for LRU behavior
    cache.delete(fingerprint);
    cache.set(fingerprint, entry);
    return entry;
}

export function setCachedEngine(fingerprint, entry) {
    if (!fingerprint || !entry) return;

    if (cache.has(fingerprint)) cache.delete(fingerprint);
    cache.set(fingerprint, entry);

    while (cache.size > MAX_CACHE_ITEMS) {
        const oldestKey = cache.keys().next().value;
        cache.delete(oldestKey);
    }
}

export function clearEngineCache() {
    cache.clear();
}
