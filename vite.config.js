import { defineConfig } from 'vite';
import fs from 'fs';
import path from 'path';
import os from 'os';

async function getHttpsOptions() {
  try {
    const certsDir = path.join(os.homedir(), '.office-addin-dev-certs');
    const keyFile = path.join(certsDir, 'localhost.key');
    const certFile = path.join(certsDir, 'localhost.crt');

    if (fs.existsSync(keyFile) && fs.existsSync(certFile)) {
      return {
        key: fs.readFileSync(keyFile),
        cert: fs.readFileSync(certFile),
      };
    }
  } catch (e) {
    console.warn("Could not load dev certs. Run 'npx office-addin-dev-certs install' first.");
  }
  return false;
}

export default async () => defineConfig({
  server: {
    port: 3000,
    https: await getHttpsOptions()
  }
});
