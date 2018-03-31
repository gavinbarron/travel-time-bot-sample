module.exports = function(config, webpack) {
    config.plugins.push(new webpack.DefinePlugin({ "global.GENTLY": false }));
    return config;
}