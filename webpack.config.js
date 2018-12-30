// @ts-check
const path = require("path");
const ExtractTextPlugin = require("extract-text-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const webpack = require("webpack");

module.exports = [
    {
        mode: "development",
        entry: {
            "app": "./src/typescript/app.ts",
        },
        output: {
            path: path.join(__dirname, "dist"),
            filename: "[name].js",
        },
        devtool: "source-map",
        resolve: {
            modules: [
                path.resolve("./src/typescript"),
                path.resolve("./node_modules")
            ],
            extensions: [".js", ".ts"],
            alias: {
                "xlsx-populate": "xlsx-populate/browser/xlsx-populate.min.js"
            }
        },
        module: {
            rules: [{
                test: /\.ts?$/,
                use: {
                    loader: "ts-loader",
                    options: {
                        compilerOptions: {
                            "noUnusedLocals": false,
                            "noUnusedParameters": false,
                        }
                    }
                }
            }]
        },
        target: "web",
        node: {
             fs: "empty",
        }
    },
    // sass -> css build
    {
        mode: "development",
        context: path.join(__dirname, "src/scss"),
        entry: {
            style: "./main.scss"
        },
        output: {
            path: path.join(__dirname, "dist"),
            filename: "style.css"
        },
        module: {
            rules: [{
                test: /\.scss$/,
                use: ExtractTextPlugin.extract({
                    use: [{
                        loader: "css-loader",
                        options: {
                            minimize: true,
                            sourceMap: true
                        }
                    }, {
                        loader: "sass-loader",
                        options: {
                            sourceMap: true
                        }
                    }]
                })
            }, {
                test: /\.png$/,
                use: [{
                    loader: "url-loader",
                    options: {
                        limit: 10000,
                        mimetype: "image/png"
                    }
                }]
            }, {
                test: /\.svg$/,
                use: [{
                    loader: "svg-url-loader",
                    options: {
                        limit: 10000,
                    }
                }]
            }]
        },
        devtool: "source-map",
        plugins: [
            new ExtractTextPlugin({
                filename: "style.css"
            }),
            new CopyWebpackPlugin([{
                from: path.join(__dirname, "src/static"),
            }])
        ]
    }
];