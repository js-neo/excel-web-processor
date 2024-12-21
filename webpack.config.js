import path from 'path';

export default {
    entry: './src/renderer.js',
    output: {
        filename: 'bundle.js',
        path: path.resolve(process.cwd(), 'dist'),
    },
    mode: 'development',
};
