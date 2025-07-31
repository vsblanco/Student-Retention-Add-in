/*
 * This file is the main entry point for all ribbon button commands.
 * It imports functions from other modules and associates them with Office actions.
 */
import { openImportDialog } from './import.js';
import { openCreateLdaDialog } from './lda.js';
import { toggleHighlight, transferData } from './actions.js';

// Register ribbon button commands
Office.actions.associate("toggleHighlight", toggleHighlight);
Office.actions.associate("openImportDialog", openImportDialog);
Office.actions.associate("transferData", transferData);
Office.actions.associate("openCreateLdaDialog", openCreateLdaDialog);
