import path from 'path';
import HtmlWebpackPlugin from 'html-webpack-plugin';
import { CleanWebpackPlugin } from 'clean-webpack-plugin';

const isDev = process.env.NODE_ENV === 'development';

export default {
    entry: './src/renderer.js',
    output: {
        filename: 'bundle.js',
        path: path.resolve(process.cwd(), 'dist'),
        publicPath: isDev ? '/' : '/excel-web-processor/',
    },
    mode: 'development',
    devtool: 'inline-source-map',
    devServer: {
        static: {
            directory: path.join(process.cwd(), 'dist'),
        },
        hot: true,
        open: true,
        client: {
            overlay: {
                errors: true,
                warnings: false,
            },
        },
        devMiddleware: {
            writeToDisk: true,
        },
    },
    plugins: [
        new CleanWebpackPlugin({
            cleanStaleWebpackAssets: false,
        }),
        new HtmlWebpackPlugin({
            template: './index.html',
            filename: 'index.html',
            inject: 'body',
            scriptLoading: 'defer',
        }),
    ],
    module: {
        rules: [
            {
                test: /\.css$/,
                use: [
                    'style-loader',
                    {
                        loader: 'css-loader',
                        options: {
                            esModule: false,
                        },
                    },
                ],
            },
            {
                test: /\.js$/,
                exclude: /node_modules/,
                use: {
                    loader: 'babel-loader',
                    options: {
                        presets: ['@babel/preset-env'],
                    },
                },
            },
        ],
    },
};