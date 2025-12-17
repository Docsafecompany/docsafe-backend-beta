// lib/policy/riskScore.js

export function getRiskLevel(score) {
  if (score >= 90) return "safe";
  if (score >= 70) return "low";
  if (score >= 50) return "medium";
  if (score >= 25) return "high";
  return "critical";
}

export function calculateRiskScore(summary, detections = null) {
  let score = 100;
  const breakdown = {};

  if (summary.critical > 0) {
    const penalty = summary.critical * 25;
    score -= penalty;
    breakdown.critical = penalty;
  }
  if (summary.high > 0) {
    const penalty = summary.high * 10;
    score -= penalty;
    breakdown.high = penalty;
  }
  if (summary.medium > 0) {
    const penalty = summary.medium * 5;
    score -= penalty;
    breakdown.medium = penalty;
  }
  if (summary.low > 0) {
    const penalty = summary.low * 2;
    score -= penalty;
    breakdown.low = penalty;
  }

  if (detections) {
    const sensitiveCount = detections.sensitiveData?.length || 0;
    if (sensitiveCount > 0) {
      const penalty = Math.min(sensitiveCount * 25, 50);
      score -= penalty;
      breakdown.sensitiveData = penalty;
    }

    const macrosCount = detections.macros?.length || 0;
    if (macrosCount > 0) {
      const penalty = Math.min(macrosCount * 15, 30);
      score -= penalty;
      breakdown.macros = penalty;
    }

    const hiddenCount = (detections.hiddenContent?.length || 0) + (detections.hiddenSheets?.length || 0);
    if (hiddenCount > 0) {
      const penalty = Math.min(hiddenCount * 8, 24);
      score -= penalty;
      breakdown.hiddenContent = penalty;
    }

    const commentsCount = detections.comments?.length || 0;
    if (commentsCount > 0) {
      const penalty = Math.min(commentsCount * 3, 15);
      score -= penalty;
      breakdown.comments = penalty;
    }

    const trackChangesCount = detections.trackChanges?.length || 0;
    if (trackChangesCount > 0) {
      const penalty = Math.min(trackChangesCount * 3, 15);
      score -= penalty;
      breakdown.trackChanges = penalty;
    }

    const metadataCount = detections.metadata?.length || 0;
    if (metadataCount > 0) {
      const penalty = Math.min(metadataCount * 2, 10);
      score -= penalty;
      breakdown.metadata = penalty;
    }

    const embeddedCount = detections.embeddedObjects?.length || 0;
    if (embeddedCount > 0) {
      const penalty = Math.min(embeddedCount * 5, 15);
      score -= penalty;
      breakdown.embeddedObjects = penalty;
    }

    const spellingCount = detections.spellingErrors?.length || 0;
    if (spellingCount > 0) {
      const penalty = Math.min(spellingCount * 1, 10);
      score -= penalty;
      breakdown.spellingGrammar = penalty;
    }

    const brokenLinksCount = detections.brokenLinks?.length || 0;
    if (brokenLinksCount > 0) {
      const penalty = Math.min(brokenLinksCount * 4, 12);
      score -= penalty;
      breakdown.brokenLinks = penalty;
    }

    const complianceCount = detections.complianceRisks?.length || 0;
    if (complianceCount > 0) {
      const penalty = Math.min(complianceCount * 12, 36);
      score -= penalty;
      breakdown.complianceRisks = penalty;
    }
  } else {
    const classifiedIssues = (summary.critical || 0) + (summary.high || 0) + (summary.medium || 0) + (summary.low || 0);
    const unclassifiedIssues = Math.max(0, (summary.totalIssues || 0) - classifiedIssues);

    if (unclassifiedIssues > 0) {
      const penalty = unclassifiedIssues * 5;
      score -= penalty;
      breakdown.unclassified = penalty;
    }
  }

  if (summary.totalIssues > 10) {
    const penalty = (summary.totalIssues - 10) * 2;
    score -= penalty;
    breakdown.volumePenalty = penalty;
  }

  const finalScore = Math.max(0, Math.min(100, score));
  return { score: finalScore, breakdown };
}

export function calculateAfterScore(beforeScore, cleaningStats, correctionStats, riskBreakdown = {}, extraRemovals = {}) {
  if (beforeScore === null || beforeScore === undefined) {
    return { score: null, scoreImpacts: {}, improvement: 0 };
  }

  let improvement = 0;
  const scoreImpacts = {};

  if (cleaningStats?.metaRemoved > 0) {
    const impact = Math.min(cleaningStats.metaRemoved * 2, riskBreakdown.metadata ?? 10);
    improvement += impact;
    scoreImpacts.metadata = impact;
  }

  if (cleaningStats?.commentsXmlRemoved > 0) {
    const impact = Math.min(cleaningStats.commentsXmlRemoved * 3, riskBreakdown.comments ?? 15);
    improvement += impact;
    scoreImpacts.comments = impact;
  }

  const trackChangesTotal =
    (cleaningStats?.revisionsAccepted?.deletionsRemoved || 0) +
    (cleaningStats?.revisionsAccepted?.insertionsUnwrapped || 0);

  if (trackChangesTotal > 0) {
    const impact = Math.min(trackChangesTotal * 3, riskBreakdown.trackChanges ?? 15);
    improvement += impact;
    scoreImpacts.trackChanges = impact;
  }

  if (cleaningStats?.hiddenRemoved > 0) {
    const impact = Math.min(cleaningStats.hiddenRemoved * 8, riskBreakdown.hiddenContent ?? 24);
    improvement += impact;
    scoreImpacts.hiddenContent = impact;
  }

  if (cleaningStats?.macrosRemoved > 0) {
    const impact = Math.min(cleaningStats.macrosRemoved * 15, riskBreakdown.macros ?? 30);
    improvement += impact;
    scoreImpacts.macros = impact;
  }

  const embeddedTotal = (cleaningStats?.mediaDeleted || 0) + (cleaningStats?.picturesRemoved || 0);
  if (embeddedTotal > 0) {
    const impact = Math.min(embeddedTotal * 5, riskBreakdown.embeddedObjects ?? 15);
    improvement += impact;
    scoreImpacts.embeddedObjects = impact;
  }

  if (correctionStats?.changedTextNodes > 0) {
    const impact = Math.min(correctionStats.changedTextNodes * 1, riskBreakdown.spellingGrammar ?? 10);
    improvement += impact;
    scoreImpacts.spellingGrammar = impact;
  }

  if (extraRemovals.sensitiveDataRemoved > 0) {
    const impact = Math.min(extraRemovals.sensitiveDataRemoved * 25, riskBreakdown.sensitiveData ?? 50);
    improvement += impact;
    scoreImpacts.sensitiveData = impact;
  }

  if (extraRemovals.hiddenContentRemoved > 0) {
    const impact = Math.min(extraRemovals.hiddenContentRemoved * 8, riskBreakdown.hiddenContent ?? 24);
    improvement += impact;
    scoreImpacts.hiddenContentExtra = impact;
  }

  const afterScore = Math.min(100, beforeScore + improvement);
  return { score: afterScore, scoreImpacts, improvement };
}
