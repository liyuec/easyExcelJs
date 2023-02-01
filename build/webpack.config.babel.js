const path = require('path');
const webpack = require('webpack');
const {CleanWebpackPlugin} = require('clean-webpack-plugin');

/*
* @type {import('webpack').Configuration}  
*/
export default {
    mode:'production',
    entry: path.resolve(__dirname,'../src/main.js'),
    output:{
        path: path.resolve(__dirname, '../dist'),
        filename: 'index.umd.min.js',
        libraryTarget:'umd',
        globalObject:'this',
        library:'easyExcel'
    },
    devtool:'source-map',
    optimization:{
        minimize:false
        //minimize:true
    },
    externals: ['exceljs','file-saver'],
    module:{
        rules:[
            {
                test: /\.(js)$/,
                //exclude: /(node_modules|bower_components)/,
                exclude:/node_modules/,
                use: 'babel-loader'
            }
        ]
    },
    plugins:[
        new CleanWebpackPlugin()
    ]
}