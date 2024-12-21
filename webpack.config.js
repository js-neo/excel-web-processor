import path from 'path';
import HtmlWebpackPlugin from 'html-webpack-plugin';

export default {
    entry: './src/renderer.js',
    output: {
        filename: 'bundle.js',
        path: path.resolve(process.cwd(), 'dist'),
        publicPath: '/excel-web-processor/',
    },
    mode: 'development',
    plugins: [
        new HtmlWebpackPlugin({
            template: './index.html',
            filename: 'index.html',
            inject: 'body',
        }),
    ],
};


