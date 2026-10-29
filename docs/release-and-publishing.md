# Release and publishing

This document records the release procedure. Each release must be validated and explicitly authorized before publication.

## Current package metadata

- Name: `@dukeofwest/universal-file-preview`
- Current version: `0.1.1`
- npm status: `0.1.1` is published publicly at [@dukeofwest/universal-file-preview](https://www.npmjs.com/package/@dukeofwest/universal-file-preview) (verified on 2026-10-29).
- Install: `npm install @dukeofwest/universal-file-preview`
- Module format: ESM only; no CommonJS export is configured.
- License metadata: MIT, matching the existing `LICENSE` file.
- Runtime dependencies: none.
- GitHub: [dukeofwest/universal-file-preview](https://github.com/dukeofwest/universal-file-preview), primary branch `main`.

The MIT license and copyright notice were present in the project before release preparation. For future releases, verify that all included code is properly licensed and attributed.

## Before a GitHub release

1. Review all source, documentation, generated output policy, secrets, and license/attribution. Preserve the existing package lockfile.
2. Keep `.env` files, credential-bearing `.npmrc` files, tokens, private keys, local caches, and unnecessary build output out of source control.
3. Run `npm ci`, `npm test`, `npm run typecheck`, and `npm pack --dry-run`.
4. Review the exact files to commit. Do not force-push or overwrite a remote.
5. Confirm CI on `main` before tagging a release.

Verify the repository URL, pushed branch, commit, and CI result with GitHub before reporting a release.

## Before npm publication

1. Confirm that the package name is available to and controlled by the intended npm account. The 404 result is not sufficient proof of availability.
2. Verify secure npm authentication and permission without sharing a token in chat or source control.
3. The current published version is `0.1.1`; use `0.1.2` for the next patch release. Never reuse a published version.
4. Ensure the reviewed source changes are committed and the working tree is clean.
5. Run the build, tests, typecheck, and package dry run; inspect the tarball contents and secrets.
6. Install the actual tarball in a clean consumer project and verify package exports, declarations, CSS subpath, and documented examples.
7. Obtain explicit publication authorization, then publish the scoped package with `npm publish --access public`.
8. Verify the registry version and repeat the clean-consumer install using the published package. Update the changelog and installation status only after registry confirmation.

Do not configure automatic npm publishing until account permissions, trusted publishing, and release protections have been confirmed. CI should test builds and pull requests but must not publish packages automatically at this stage.

## Next patch release command checklist

The following is a reference checklist for a future `0.1.2` release. Review its output and obtain explicit maintainer approval before running the commit, push, tag, or publish commands.

```sh
npm version patch --no-git-tag-version
npm ci
npm run build
npm test
npm run typecheck
npm pack --dry-run
npm pack
npm install --prefix ../ufp-consumer ./dukeofwest-universal-file-preview-0.1.2.tgz
```

Initialize `../ufp-consumer` as a clean npm project before the local tarball install. In that consumer directory, verify the root ESM import, generated TypeScript declarations, and CSS subpath with the consumer's bundler. `npm pack --dry-run` alone does not verify imports.

After updating the changelog for `0.1.2` and reviewing the exact file list, the following commands show the release sequence. Inspect `git status` and `git log origin/main..HEAD` before pushing because a push of `main` sends every commit in that range.

```sh
git status --short --branch
git diff --check
git add package.json package-lock.json README.md CHANGELOG.md SECURITY.md apps/demo/server.mjs .github docs
git commit -m "chore: prepare 0.1.2 release"
git tag v0.1.2
git push origin main v0.1.2
```

Once that remote change is approved and verified, publish with:

```sh
npm publish --access public
```

Do not paste npm tokens into chat, files, or command history. Authenticate through npm's supported login or trusted publishing flow.

## Versioning and release notes

Use semantic versioning: patch for compatible fixes, minor for compatible features, and major for public API breaks. Add a dated changelog entry describing user-visible changes and known limits. Tag only a validated, committed release. Do not claim a push, tag, or npm publication until the relevant service confirms it.
