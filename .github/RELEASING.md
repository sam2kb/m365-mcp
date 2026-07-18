# Releasing

The `CI and Release` workflow tests and packages every pull request and push to
`main`. Tags matching `v*` publish the same version to npm and ClawHub.

## One-time setup

### npm trusted publisher

In the npm settings for `@sam2kb/m365-mcp`, add a GitHub Actions trusted
publisher with:

- Organization or user: `sam2kb`
- Repository: `m365-mcp`
- Workflow filename: `release.yml`
- Allowed action: `npm publish`
- Environment: leave unset

The workflow uses npm OIDC trusted publishing, so it does not require an
`NPM_TOKEN` secret.

### ClawHub token

Create a ClawHub token for `sam2kb` with permission to publish `m365-mcp`,
then add it to the GitHub repository as an Actions secret named
`CLAWHUB_TOKEN`.

ClawHub skill publishing does not currently support OIDC, so this secret is
required for tagged releases.

## Publish a release

Start from a clean `main` branch:

```bash
npm version patch
git push origin main --follow-tags
```

Use `minor` or `major` instead of `patch` when appropriate. The generated
tag must match the version in `package.json`; the workflow stops before
publishing if they differ.

The release runs in this order:

1. Install dependencies, run all tests, and build.
2. Inspect the npm tarball.
3. Publish the prebuilt package to npm using OIDC.
4. Publish the npm-backed skill to ClawHub using `CLAWHUB_TOKEN`.
