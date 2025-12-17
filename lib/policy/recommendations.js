// lib/policy/recommendations.js

export function generateRecommendations(detections) {
  const recommendations = [];

  if (detections.metadata?.length > 0)
    recommendations.push("Remove document metadata to protect author and organization information.");
  if (detections.comments?.length > 0)
    recommendations.push(`Review and remove ${detections.comments.length} comment(s) before sharing externally.`);
  if (detections.trackChanges?.length > 0)
    recommendations.push("Accept or reject all tracked changes to finalize the document.");
  if (detections.hiddenContent?.length > 0 || detections.hiddenSheets?.length > 0)
    recommendations.push("Remove hidden content that could expose confidential information.");
  if (detections.macros?.length > 0)
    recommendations.push("Remove macros for security - they can contain executable code.");
  if (detections.sensitiveData?.length > 0) {
    const types = [...new Set(detections.sensitiveData.map((d) => d.type))];
    recommendations.push(`Review sensitive data detected: ${types.join(", ")}.`);
  }
  if (detections.embeddedObjects?.length > 0)
    recommendations.push("Remove embedded objects that may contain hidden data.");
  if (detections.spellingErrors?.length > 0)
    recommendations.push(`${detections.spellingErrors.length} spelling/grammar issue(s) were detected.`);

  if (recommendations.length === 0)
    recommendations.push("Document appears clean. Minor review recommended before external sharing.");

  return recommendations;
}
