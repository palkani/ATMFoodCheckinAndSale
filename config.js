import path from 'node:path';
import { fileURLToPath } from 'node:url';

// Where the database, Excel files and uploaded pictures live. In the cloud this must be a mounted persistent volume.
export const DATA_DIR = process.env.DATA_DIR || path.join(path.dirname(fileURLToPath(import.meta.url)), 'data');
