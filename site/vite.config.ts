import { sveltekit } from '@sveltejs/kit/vite';
import { defineConfig } from 'vite';

export default defineConfig({
  plugins: [sveltekit()],
  server: {
    fs: {
      // Allow serving files from the parent project (for example .ts files
      // imported via ?raw and the xlsx-kit source via path alias).
      allow: ['..'],
    },
  },
});
