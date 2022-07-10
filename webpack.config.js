const path = require("path");
const GasPlugin = require("gas-webpack-plugin");

module.exports = {
  mode: "production",
  entry: "./src/index.ts",
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
    ],
  },
  resolve: {
    extensions: [".tsx", ".ts", ".js"],
  },
  output: {
    filename: "Code.gs",
    path: path.resolve(__dirname, "dist"),
    clean: true,
  },
  optimization: {
    minimize: false,
  },
  plugins: [new GasPlugin()],
};
