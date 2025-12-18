import { checkSpellingWithAI as spellcheck } from "./aiSpellCheck.js";
import { microDeterministicIssues } from "./aiProofreadAnchored.js";

function normType(t) {
  const x = String(t || "").toLowerCase();
  if (x.includes("multiple")) return "multiple_spaces";
  if (x.includes("fragment")) return "fragmented_word";
  if (x === "spacing") return "spacing";
  return x || "spelling";
}

function toUnified(item, source) {
  return {
    id: item.id || `${source}_${Math.random().toString(16).slice(2)}`,
    error: item.error,
    correction: item.correction,
    type: normType(item.type),
    severity: item.severity || "medium",
    message: item.message || "",
    globalStart: Number.isFinite(item.globalStart) ? item.globalStart : null,
    globalEnd: Number.isFinite(item.globalEnd) ? item.globalEnd : null,
    source
  };
}

function overlaps(a, b) {
  if (a.globalStart == null || a.globalEnd == null) return false;
  if (b.globalStart == null || b.globalEnd == null) return false;
  return a.globalStart < b.globalEnd && b.globalStart < a.globalEnd;
}

/**
 * Keep best when overlaps:
 * - Prefer aiSpellCheck (offset AI / candidate-validated)
 * - If same source, prefer shorter span (safer)
 */
function pickBest(existing, candidate) {
  if (!existing) return candidate;
  if (existing.source !== candidate.source) {
    if (existing.source === "aiSpellCheck") return existing;
    if (candidate.source === "aiSpellCheck") return candidate;
  }
  const lenE = existing.globalEnd - existing.globalStart;
  const lenC = candidate.globalEnd - candidate.globalStart;
  return lenC < lenE ? candidate : existing;
}

function dedupAndResolve(list) {
  // 1) hard dedup by exact span + correction
  const byKey = new Map();
  for (const x of list) {
    if (x.globalStart == null || x.globalEnd == null) continue;
    const key = `${x.globalStart}:${x.globalEnd}:${String(x.correction||"").toLowerCase()}`;
    const prev = byKey.get(key);
    byKey.set(key, prev ? pickBest(prev, x) : x);
  }

  // 2) sort and resolve overlaps
  const sorted = Array.from(byKey.values()).sort((a, b) => a.globalStart - b.globalStart);
  const final = [];
  for (const x of sorted) {
    const last = final[final.length - 1];
    if (!last) { final.push(x); continue; }

    if (!overlaps(last, x)) {
      final.push(x);
    } else {
      // overlap => keep the best one
      final[final.length - 1] = pickBest(last, x);
    }
  }
  return final;
}

/**
 * Combined proofread:
 * - Primary: aiSpellCheck (high recall + safe offsets)
 * - Secondary: microDeterministic short-fragment fixes from anchored module
 */
export async function checkSpellingCombined(text) {
  const t = String(text || "");
  if (t.trim().length < 10) return [];

  const main = await spellcheck(t); // returns items with globalStart/globalEnd
  const micro = microDeterministicIssues(t);

  const unified = [
    ...main.map(x => toUnified(x, "aiSpellCheck")),
    ...micro.map(x => toUnified(x, "microDet")),
  ].filter(x => x.globalStart != null && x.globalEnd != null);

  return dedupAndResolve(unified);
}

export default { checkSpellingCombined };
