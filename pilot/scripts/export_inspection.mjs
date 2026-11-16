import fs from 'node:fs/promises';

// Box and other mounted folders can cross filesystem boundaries even under C:.
export async function relocateExportInspection(source, destination, files = fs) {
  try {
    await files.rename(source, destination);
  } catch (error) {
    if (error.code === 'ENOENT') return; // Some runtime versions emit no sidecar.
    if (error.code !== 'EXDEV') throw error;
    // Remove the original only after the diagnostic has been copied successfully.
    await files.copyFile(source, destination);
    await files.unlink(source);
  }
}
