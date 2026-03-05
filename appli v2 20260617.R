// Updated session loading logic to ensure that incomplete entries remain incomplete during restore.

function restoreSession(entries) {
    entries.forEach(entry => {
        if (entry.isComplete) {
            // Apply original logic for complete entries
            fallback_bd1(entry);
        } else {
            // For truly incomplete entries
            entry.haz_output = null; // Ensure haz_output is set to null for incomplete entries
        }
    });
}