import JSZip from 'jszip';
import { normalizeText } from './textCleaner.js';

// Nettoie DOCX en conservant la mise en forme : on traite chaque <w:t>
export async function cleanDOCX(buffer) {
  const zip = await JSZip.loadAsync(buffer);

  // Retirer métadonnées/commentaires
  const remove = [
    'docProps/core.xml',
    'docProps/app.xml',
    'docProps/custom.xml',
    '
