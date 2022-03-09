const path = require('path');

module.exports = {
    mode: 'development',
    entry: path.join(__dirname, 'src/index.js'),
    output: {
        path: path.join(__dirname, './build'),
        filename: 'bundle.js',
        globalObject: 'this'
    },
    externals: ['fs', 'os', 'path'],
    module: {
        rules: [
            {
                test: /\.(js)$/,
                exclude: /node_modules/,
                use: ['babel-loader']
            }
        ]
    }
}