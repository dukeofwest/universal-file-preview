# Release and publishing

This document records the release procedure. Each release must be validated and explicitly authorized before publication.

## Current package metadata

- Name: `@dukeofwest/universal-file-preview`
- Version: `0.1.0`
- npm status: published publicly at [@dukeofwest/universal-file-preview](https://www.npmjs.com/package/@dukeofwest/universal-file-preview).
- Install: `npm install @dukeofwest/universal-file-preview`
- Module format: ESM only; no CommonJS export is configured.
- License metadata: MIT, matching the existing `LICENSE` file.
- Runtime dependencies: none.
- GitHub: [dukeofwest/universal-file-preview](https://github.com/dukeofwest/universal-file-preview), primary branch `main`.

The MIT license and copyright notice were present in the project before release preparation. For future releases, verify that all included code is properly licensed and attributed.

## Before a GitHub release

1. Confirm the GitHub account and exact repository URL; do not infer an organization or username.
2. Review all source, documentation, generated output policy, secrets, and license/attribution. Preserve the existing package lockfile.
3. Keep `.env` files, credential-bearing `.npmrc` files, tokens, private keys, local caches, and unnecessary build output out of source control.
4. Run `npm ci`, `npm test`, `npm run typecheck`, and `npm pack --dry-run`.
5. Review the exact files to commit. Do not force-push or overwrite a remote.
6. Confirm CI on the intended primary branch before tagging a release.

Verify the repository URL, pushed branch, commit, and CI result with GitHub before reporting a release.

## Before npm publication

1. Confirm that the package name is available to and controlled by the intended npm account. The 404 result is not sufficient proof of availability.
2. Verify secure npm authentication and permission without sharing a token in chat or source control.
3. Decide the release version and public package access. Version `0.1.0` is currently configured; do not reuse it if that version exists at publication time.
4. Ensure the reviewed source changes are committed and the working tree is clean.
5. Run the build, tests, typecheck, and package dry run; inspect the tarball contents and secrets.
6. Install the actual tarball in a clean consumer project and verify package exports, declarations, CSS subpath, and documented examples.
7. Obtain explicit publication authorization, then publish the scoped package with `npm publish --access public`.
8. Verify the registry version and repeat the clean-consumer install using the published package. Update the changelog and installation status only after registry confirmation.

Do not configure automatic npm publishing until account permissions, trusted publishing, and release protections have been confirmed. CI should test builds and pull requests but must not publish packages automatically at this stage.

## Versioning and release notes

Use semantic versioning: patch for compatible fixes, minor for compatible features, and major for public API breaks. Add a dated changelog entry describing user-visible changes and known limits. Tag only a validated, committed release. Do not claim a push, tag, or npm publication until the relevant service confirms it.
