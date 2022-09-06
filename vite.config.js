import { fileURLToPath, URL } from 'node:url';

import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';

// https://vitejs.dev/config/
export default defineConfig({
    base: '',
    plugins: [vue()],
    resolve: {
        alias: {
            '@': fileURLToPath(new URL('./src', import.meta.url)),
            'crypto': `crypto-browserify`,
            'buffer': '@red0531/buffer'
        }
    },
    define: {
        global: 'globalThis',
        'process.env': '{}'
    },

    build: {
        target: [ 'es2020' ]
    },

    optimizeDeps: {
        esbuildOptions: {
            target: 'esnext',
        },
    },
})
