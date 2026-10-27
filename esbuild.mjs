import esbuild from 'esbuild';
await esbuild.build({entryPoints:['src/extension.ts'],outfile:'dist/extension.js',bundle:true,platform:'node',target:'node20',format:'cjs',external:['vscode'],logLevel:'info'});
await esbuild.build({entryPoints:['src/workbook.ts'],outfile:'dist/workbook.js',bundle:true,platform:'node',target:'node20',format:'cjs',logLevel:'info'});
