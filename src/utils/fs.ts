import { promises as fs } from 'fs';
import * as path from 'path';

export async function writeFileIfChanged(filePath: string, newContent: string): Promise<void> {
  let existingContent: string | undefined;
  try {
    existingContent = await fs.readFile(filePath, 'utf8');
  } catch (error) {
    const nodeError = error as NodeJS.ErrnoException;
    if (nodeError.code && nodeError.code !== 'ENOENT') {
      throw error;
    }
  }

  if (existingContent === newContent) {
    return;
  }

  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, newContent, 'utf8');
}
