// lib/cache.js
// VERSION 1.0 - Intelligent caching for document analysis

import crypto from 'crypto';

// In-memory cache (use Redis for production persistence)
const analysisCache = new Map();
const CACHE_TTL_MS = 30 * 60 * 1000; // 30 minutes
const MAX_CACHE_ENTRIES = 100;

/**
 * Generate SHA-256 hash of file content
 */
export function getFileHash(buffer) {
  return crypto.createHash('sha256').update(buffer).digest('hex');
}

/**
 * Get cached analysis result if available and not expired
 */
export function getCachedAnalysis(fileHash) {
  const cached = analysisCache.get(fileHash);
  
  if (!cached) {
    return null;
  }
  
  // Check if expired
  if (Date.now() - cached.timestamp > CACHE_TTL_MS) {
    analysisCache.delete(fileHash);
    console.log(`[CACHE] Expired entry removed: ${fileHash.slice(0, 8)}...`);
    return null;
  }
  
  console.log(`[CACHE] HIT for ${fileHash.slice(0, 8)}... (age: ${Math.round((Date.now() - cached.timestamp) / 1000)}s)`);
  return cached.data;
}

/**
 * Store analysis result in cache
 */
export function setCachedAnalysis(fileHash, data) {
  // Clean up old entries if at max capacity
  if (analysisCache.size >= MAX_CACHE_ENTRIES) {
    const entries = [...analysisCache.entries()];
    entries.sort((a, b) => a[1].timestamp - b[1].timestamp);
    
    // Remove oldest 10%
    const toRemove = Math.ceil(MAX_CACHE_ENTRIES * 0.1);
    for (let i = 0; i < toRemove; i++) {
      analysisCache.delete(entries[i][0]);
    }
    console.log(`[CACHE] Cleaned up ${toRemove} old entries`);
  }
  
  analysisCache.set(fileHash, {
    data,
    timestamp: Date.now()
  });
  
  console.log(`[CACHE] Stored analysis for ${fileHash.slice(0, 8)}... (total entries: ${analysisCache.size})`);
}

/**
 * Clear all cache entries
 */
export function clearCache() {
  const size = analysisCache.size;
  analysisCache.clear();
  console.log(`[CACHE] Cleared ${size} entries`);
}

/**
 * Get cache statistics
 */
export function getCacheStats() {
  return {
    entries: analysisCache.size,
    maxEntries: MAX_CACHE_ENTRIES,
    ttlMinutes: CACHE_TTL_MS / 60000
  };
}

export default { getFileHash, getCachedAnalysis, setCachedAnalysis, clearCache, getCacheStats };
