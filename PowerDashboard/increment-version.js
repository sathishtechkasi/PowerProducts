const fs = require('fs');
const path = require('path');

// Target the package-solution.json file
const packageSolutionPath = path.join(__dirname, 'config', 'package-solution.json');

try {
  if (fs.existsSync(packageSolutionPath)) {
    const packageSolution = JSON.parse(fs.readFileSync(packageSolutionPath, 'utf8'));
    const currentVersion = packageSolution.solution.version; // e.g., "1.0.0.0"

    // Split the version into its 4 parts
    const versionParts = currentVersion.split('.');
    
    if (versionParts.length === 4) {
      // Increment the 4th digit (the revision/patch number)
      let revision = parseInt(versionParts[3], 10);
      revision += 1;
      versionParts[3] = revision.toString();

      // Join it back together
      const newVersion = versionParts.join('.');
      packageSolution.solution.version = newVersion;

      // Write the updated file back to the disk
      fs.writeFileSync(packageSolutionPath, JSON.stringify(packageSolution, null, 2), 'utf8');
      console.log(`\n🚀 \x1b[32mSPFx Version automatically bumped: ${currentVersion} -> ${newVersion}\x1b[0m\n`);
    } else {
      console.error('\x1b[31mError: package-solution.json version must be in X.X.X.X format.\x1b[0m');
    }
  } else {
    console.error('\x1b[31mError: config/package-solution.json not found!\x1b[0m');
  }
} catch (err) {
  console.error('\x1b[31mError bumping version:\x1b[0m', err.message);
}