/* eslint-disable no-undef */  

const devCerts = require("office-addin-dev-certs");  
const CopyWebpackPlugin = require("copy-webpack-plugin");  
const HtmlWebpackPlugin = require("html-webpack-plugin");  
const path = require("path");  

const urlDev = "https://localhost:3000/";  
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION  

async function getHttpsOptions() {  
  const httpsOptions = await devCerts.getHttpsServerOptions();  
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };  
}  

module.exports = async (env, options) => {  
  const dev = options.mode === "development";  
  const httpsOptions = await getHttpsOptions(); // Get HTTPS options  

  const config = {  
    devtool: "source-map",  
    entry: {  
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],  
      taskpane: "./src/taskpane/taskpane.js",  
      commands: "./src/commands/commands.js",  
    },  
    output: {  
      clean: true,  
      filename: "[name].bundle.js",  
      path: path.resolve(__dirname, "dist"),  
      publicPath: "/", // Ensure public path is set correctly  
    },  
    resolve: {  
      extensions: [".html", ".js"],  
    },  
    module: {  
      rules: [  
        {  
          test: /\.js$/,  
          exclude: /node_modules/,  
          use: {  
            loader: "babel-loader",  
          },  
        },  
        {  
          test: /\.html$/,  
          exclude: /node_modules/,  
          use: "html-loader",  
        },  
        {  
          test: /\.(png|jpg|jpeg|gif|ico)$/,  
          type: "asset/resource",  
          generator: {  
            filename: "assets/[name][ext][query]",  
          },  
        },  
      ],  
    },  
    plugins: [  
      new HtmlWebpackPlugin({  
        filename: "taskpane.html",  
        template: "./src/taskpane/taskpane.html",  
        chunks: ["polyfill", "taskpane"],  
      }),  
      new CopyWebpackPlugin({  
        patterns: [  
          {  
            from: "assets/*",  
            to: "assets/[name][ext][query]",  
          },  
          {  
            from: "manifest*.xml",  
            to: "[name][ext]",  
            transform(content) {  
              if (dev) {  
                return content;  
              } else {  
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);  
              }  
            },  
          },  
        ],  
      }),  
      new HtmlWebpackPlugin({  
        filename: "commands.html",  
        template: "./src/commands/commands.html",  
        chunks: ["polyfill", "commands"],  
      }),  
    ],  
    devServer: {  
      static: {  
        directory: path.resolve(__dirname, "dist"), // Correct static directory  
      },  
      historyApiFallback: true, // Ensure URLs can be handled properly  
      port: process.env.npm_package_config_dev_server_port || 3000,  
      server: {  
        type: 'https', // Enable 'https' server  
        options: {  
          key: httpsOptions.key,  
          cert: httpsOptions.cert,  
          ca: httpsOptions.ca // If you need to provide CA for your self-signed certificate   
        }  
      },  
    },  
  };  

  return config;  
};