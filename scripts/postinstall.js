const fs = require('fs');
const path = require('path');

const reactPackageJsonPath = path.join(__dirname, '../node_modules/react/package.json');
const reactDevPath = path.join(__dirname, '../node_modules/react/cjs/react.development.js');
const reactProdPath = path.join(__dirname, '../node_modules/react/cjs/react.production.min.js');

try {
  // 1. Patch Webpack ESM module exports
  if (fs.existsSync(reactPackageJsonPath)) {
    const pkg = JSON.parse(fs.readFileSync(reactPackageJsonPath, 'utf8'));
    if (!pkg.exports) {
      pkg.exports = {
        ".": "./index.js",
        "./jsx-runtime": "./jsx-runtime.js",
        "./jsx-dev-runtime": "./jsx-dev-runtime.js"
      };
      fs.writeFileSync(reactPackageJsonPath, JSON.stringify(pkg, null, 2));
      console.log('✅ Patched react/package.json for Webpack 5');
    }
  }

  // 2. Polyfill React 18 hooks for React 17
  const polyfills = `
// --- React 18 Polyfills added by postinstall ---
if (typeof exports.useInsertionEffect === 'undefined') {
  exports.useInsertionEffect = exports.useLayoutEffect || exports.useEffect;
}
if (typeof exports.useId === 'undefined') {
  var _idCounter = 0;
  exports.useId = function useId() {
    var _id = ReactSharedInternals ? ReactSharedInternals.ReactCurrentOwner && ReactSharedInternals.ReactCurrentOwner.current : null;
    return 'r-' + Math.random().toString(36).substr(2, 5) + '-' + (_idCounter++);
  };
}
if (typeof exports.useSyncExternalStore === 'undefined') {
  exports.useSyncExternalStore = function (subscribe, getSnapshot) {
    var val = getSnapshot();
    exports.useEffect(function() {
      return subscribe(function() {});
    }, [subscribe]);
    return val;
  }
}
// ---------------------------------------------
`;

  function patchFile(filePath) {
    if (fs.existsSync(filePath)) {
      let content = fs.readFileSync(filePath, 'utf8');
      if (!content.includes('useInsertionEffect = exports.useLayoutEffect')) {
        // Insert polyfill before the last closing IIFE bracket if exists, or at the end
        if (content.endsWith('  })();\n}\n')) {
          content = content.replace('  })();\n}\n', polyfills + '  })();\n}\n');
        } else {
          content += '\n' + polyfills;
        }
        fs.writeFileSync(filePath, content);
        console.log(`✅ Patched React 18 hooks into ${path.basename(filePath)}`);
      }
    }
  }

  patchFile(reactDevPath);
  patchFile(reactProdPath);
  
} catch (e) {
  console.log('Skipping React patch', e);
}
