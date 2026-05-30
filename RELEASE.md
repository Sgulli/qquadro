# Release Checklist

## Pre-Release

- [ ] All tests pass: `pnpm test`
- [ ] TypeScript builds clean: `pnpm typecheck`
- [ ] Biome linter clean: `pnpm lint`
- [ ] Benchmark passes: `pnpm bench`
- [ ] Test coverage ≥ 90%: `pnpm test -- --coverage`
- [ ] CHANGELOG.md updated with all changes since last release
- [ ] Version bumped in `packages/core/package.json`
- [ ] Version bumped in root `package.json`

## Release

- [ ] Tag the release: `git tag v$(jq -r .version packages/core/package.json)`
- [ ] Push tag: `git push --tags`
- [ ] Publish to npm: `cd packages/core && pnpm publish`
- [ ] Create GitHub Release with release notes
- [ ] Verify npm package: `npm view @qquadro/core`

## Post-Release

- [ ] Update ROADMAP.md with completed milestone
- [ ] Update README.md examples if API changed
- [ ] Announce release
