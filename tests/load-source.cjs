const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const ts = require('typescript');

// 每次加载使用独立模块缓存；只替换测试明确指定的外部边界，不改源文件。
exports.sourceLoader = function sourceLoader(overrides = {}) {
  const root = path.resolve(__dirname, '..');
  const mocks = new Map(Object.entries(overrides).map(([name, value]) => [path.resolve(root, name), value]));
  const cache = new Map();
  function load(file) {
    const filename = path.resolve(root, file);
    if (mocks.has(filename)) return mocks.get(filename);
    if (cache.has(filename)) return cache.get(filename).exports;
    const mod = new Module(filename, module);
    mod.filename = filename;
    mod.paths = Module._nodeModulePaths(path.dirname(filename));
    const nativeRequire = mod.require.bind(mod);
    mod.require = (name) => {
      if (name.startsWith('.')) {
        const target = path.resolve(path.dirname(filename), name);
        if (fs.existsSync(target + '.ts')) return load(target + '.ts');
      }
      return nativeRequire(name);
    };
    cache.set(filename, mod);
    const compiled = ts.transpileModule(fs.readFileSync(filename, 'utf8'), {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, esModuleInterop: true },
    }).outputText;
    mod._compile(compiled, filename);
    return mod.exports;
  }
  return load;
};
