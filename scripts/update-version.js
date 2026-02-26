const fs = require('fs');
const path = require('path');

// 1. package.json dosyasındaki yeni versiyonu oku (npm version komutu bunu otomatik değiştirmiş olacak)
const packageJsonPath = path.resolve(__dirname, '../package.json');
const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
const newPackageVersion = packageJson.version; // Örn: "0.0.5" veya "1.0.1"

// 2. SharePoint paketi için versiyon numarasını 4 haneli formata çevir (Örn: "1.0.1.0")
const versionSegments = newPackageVersion.split('.');
while (versionSegments.length < 4) {
  versionSegments.push('0');
}
const spfxVersion = versionSegments.slice(0, 4).join('.');

// 3. config/package-solution.json dosyasını parçala
const solutionConfigPath = path.resolve(__dirname, '../config/package-solution.json');
const solutionConfig = JSON.parse(fs.readFileSync(solutionConfigPath, 'utf8'));

// 4. İlgili versiyon kısımlarını SharePoint paketine entegre et
solutionConfig.solution.version = spfxVersion;

if (solutionConfig.solution.features && solutionConfig.solution.features.length > 0) {
  solutionConfig.solution.features.forEach(feature => {
    feature.version = spfxVersion;
  });
}

// 5. Dosyayı yeniden yaz formatlı bir şekilde kaydet
fs.writeFileSync(solutionConfigPath, JSON.stringify(solutionConfig, null, 2), 'utf8');

console.log(`✅ config/package-solution.json version '${spfxVersion}' data is persist project!`);
