const fs = require('fs');
const path = require('path');

// package.json'dan master version'u al
const packageJsonPath = path.resolve(__dirname, '../package.json');
const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
const newPackageVersion = packageJson.version;

// SPFx paketi için version numarasını 4 haneye eşitle (Örn: 1.0.1 -> 1.0.1.0)
const versionSegments = newPackageVersion.split('.');
while (versionSegments.length < 4) {
  versionSegments.push('0');
}
const spfxVersion = versionSegments.slice(0, 4).join('.');

// package-solution update
const solutionConfigPath = path.resolve(__dirname, '../config/package-solution.json');
const solutionConfig = JSON.parse(fs.readFileSync(solutionConfigPath, 'utf8'));

solutionConfig.solution.version = spfxVersion;

if (solutionConfig.solution.features && solutionConfig.solution.features.length > 0) {
  solutionConfig.solution.features.forEach(feature => {
    feature.version = spfxVersion;
  });
}

// Dosyayı kaydet
fs.writeFileSync(solutionConfigPath, JSON.stringify(solutionConfig, null, 2), 'utf8');

console.log(`✅ config/package-solution.json version '${spfxVersion}' data is persist project!`);
