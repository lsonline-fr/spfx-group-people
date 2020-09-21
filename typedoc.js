module.exports = {
    mode: "file",
    out: "./docs",
    exclude: "test,dist,node_modules",
    inputFiles: "./src",
    tsconfig: "tsconfig.json",
    readme: "README.md",
    ignoreCompilerErrors: true,
    target: "ES5"
};