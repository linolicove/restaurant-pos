import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import obfuscator from 'rollup-plugin-obfuscator'

// https://vite.dev/config/
export default defineConfig(({ mode }) => {
  const isProduction = mode === 'production';

  return {
    plugins: [
      react(),
      // Only obfuscate during production builds so local dev stays fast
      isProduction && obfuscator({
        compact: true,
        controlFlowFlattening: true,
        controlFlowFlatteningThreshold: 0.75,
        numbersToExpressions: true,
        simplify: true,
        stringArray: true,
        stringArrayThreshold: 0.8,
        splitStrings: true,
        splitStringsChunkLength: 5,
        transformObjectKeys: true,
        unicodeEscapeSequence: false
      })
    ].filter(Boolean),

    build: {
      // 1. Critical: Disables .map files so DevTools cannot inspect original source code
      sourcemap: false,

      // 2. Uses Terser for advanced minification and identifier mangling
      minify: 'terser',
      terserOptions: {
        compress: {
          drop_console: true,
          drop_debugger: true,
          pure_funcs: ['console.log', 'console.info', 'console.debug']
        },
        mangle: {
          toplevel: true
        },
        format: {
          comments: false // Strips all license headers and developer comments
        }
      },

      // 3. Scrambles chunk names so file structures are obscured
      rollupOptions: {
        output: {
          entryFileNames: 'assets/[hash].js',
          chunkFileNames: 'assets/[hash].js',
          assetFileNames: 'assets/[hash].[ext]'
        }
      }
    }
  };
});