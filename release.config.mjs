export default {
  branches: ['main'],
  tagFormat: 'v${version}',
  plugins: [
    ['@semantic-release/commit-analyzer', {
      preset: 'conventionalcommits',
      releaseRules: [{type:'ci',scope:'release',release:'patch'}],
    }],
    ['@semantic-release/release-notes-generator', {preset:'conventionalcommits'}],
    ['@semantic-release/changelog', {changelogFile:'CHANGELOG.md'}],
    ['@semantic-release/npm', {tarballDir:'dist'}],
    ['@semantic-release/exec', {prepareCmd:'node scripts/package-release.mjs ${nextRelease.version}'}],
    ['@semantic-release/git', {
      assets:['package.json','package-lock.json','CHANGELOG.md','panel/**'],
      message:'chore(release): ${nextRelease.version} [skip ci]\n\n${nextRelease.notes}',
    }],
    ['@semantic-release/github', {
      assets:['dist/*.zip','dist/*.tgz','dist/SHA256SUMS.txt'],
      successComment:false,failComment:false,failTitle:false,releasedLabels:false,
    }],
  ],
};
