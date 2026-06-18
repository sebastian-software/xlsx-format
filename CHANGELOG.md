# Changelog

## [2.3.1](https://github.com/sebastian-software/xlsx-format/compare/xlsx-format-v2.3.0...xlsx-format-v2.3.1) (2026-06-18)


### Bug Fixes

* **docs:** avoid duplicate api sidebar active state ([93202c0](https://github.com/sebastian-software/xlsx-format/commit/93202c08e91f972dfcebe3158851b0fa70b16a08))
* **docs:** neutralize ardo footer surface ([1aea064](https://github.com/sebastian-software/xlsx-format/commit/1aea0648187b2f89dfd2ea8c51ea53701865405e))
* **docs:** normalize ardo sidebar and theme ([f736f68](https://github.com/sebastian-software/xlsx-format/commit/f736f68f8c1eed83b4686e21306daaf4a39964f4))
* **docs:** restore ardo navigation and home features ([da3c8ae](https://github.com/sebastian-software/xlsx-format/commit/da3c8aee6aa3c9f63798dea669e662b8da312f4f))


### Documentation

* overhaul homepage ([a29c2fe](https://github.com/sebastian-software/xlsx-format/commit/a29c2fec3d0b3d8be21a75cab711e480c347ab7f))
* redesign homepage with a spreadsheet-native visual system ([cdbac41](https://github.com/sebastian-software/xlsx-format/commit/cdbac41adb9ac57a7caf4ae3053535267efa556c))

## [2.3.0](https://github.com/sebastian-software/xlsx-format/compare/xlsx-format-v2.2.0...xlsx-format-v2.3.0) (2026-06-18)


### Features

* secure export protections by default ([#48](https://github.com/sebastian-software/xlsx-format/issues/48)) ([7f87c60](https://github.com/sebastian-software/xlsx-format/commit/7f87c603e928aedf34995dfd5f82413b79753304))

## [2.2.0](https://github.com/sebastian-software/xlsx-format/compare/xlsx-format-v2.1.1...xlsx-format-v2.2.0) (2026-06-16)


### Features

* **api:** add typed error codes ([#46](https://github.com/sebastian-software/xlsx-format/issues/46)) ([f203e6d](https://github.com/sebastian-software/xlsx-format/commit/f203e6dae1b264ef0261b418def90861a662c977))


### Bug Fixes

* **pkg:** expose package metadata ([#44](https://github.com/sebastian-software/xlsx-format/issues/44)) ([940fa59](https://github.com/sebastian-software/xlsx-format/commit/940fa5916bf3a1a50aceeab567cf4b86caeae710))


### Documentation

* **readme:** add security section ([#45](https://github.com/sebastian-software/xlsx-format/issues/45)) ([99973aa](https://github.com/sebastian-software/xlsx-format/commit/99973aaad8329ebcb174ecc77307995f05ee6979))

## [2.1.1](https://github.com/sebastian-software/xlsx-format/compare/xlsx-format-v2.1.0...xlsx-format-v2.1.1) (2026-06-13)


### Bug Fixes

* add XML parser safety limits ([4c94073](https://github.com/sebastian-software/xlsx-format/commit/4c9407363de4ff15ccabae8449dab893a3a6ce52)), closes [#25](https://github.com/sebastian-software/xlsx-format/issues/25)

## [2.1.0](https://github.com/sebastian-software/xlsx-format/compare/xlsx-format-v2.0.0...xlsx-format-v2.1.0) (2026-06-12)


### Features

* add styled xlsx report writing ([d24aa64](https://github.com/sebastian-software/xlsx-format/commit/d24aa64c6e08d38dbaf96ee9216505c8a546a934))
* **api:** add sheetToArray and typed cell helpers ([ef1b2c8](https://github.com/sebastian-software/xlsx-format/commit/ef1b2c8fa8314630368d96c3148de854ca4978eb))


### Bug Fixes

* **api:** harden spreadsheet exports against malicious data ([1fbce35](https://github.com/sebastian-software/xlsx-format/commit/1fbce35abef484d2e446265eb865711737e61aed))
* synchronize exported package version ([b38259d](https://github.com/sebastian-software/xlsx-format/commit/b38259d2a79895d0c8423a07b99e87335ceca76b))
* updated Ardo ([dbc0126](https://github.com/sebastian-software/xlsx-format/commit/dbc0126ba5463810c2a6c5745c186c5490a8f90a))
* updated deps ([cba8dd7](https://github.com/sebastian-software/xlsx-format/commit/cba8dd7c0ff310c0826669c4655465ee9b51b440))
* **zip:** harden parsing against malicious archives ([ff2a1d3](https://github.com/sebastian-software/xlsx-format/commit/ff2a1d333188a63c9cf09acaa173085737dc4d08))


### Documentation

* add Sebastian Software branding to README ([68d7661](https://github.com/sebastian-software/xlsx-format/commit/68d7661acf90b70b1e6c8ba40133244f21ac8f2a))
* add Sebastian Software README branding ([e130c84](https://github.com/sebastian-software/xlsx-format/commit/e130c84841cf4d47d941f5e8d3c50c93a9a003bc))
* **security:** add policy and guidance ([3d09f8b](https://github.com/sebastian-software/xlsx-format/commit/3d09f8b9fbebf40113e3c2ef69361f039a96f2f5))

## [2.0.0](https://github.com/sebastian-software/xlsx-format/compare/xlsx-format-v1.0.2...xlsx-format-v2.0.0) (2026-02-12)


### ⚠ BREAKING CHANGES

* readFile and writeFile are no longer exported. Use `read(await readFile(path))` and `await writeFile(path, await write(wb))` from node:fs/promises instead.

### Features

* remove readFile/writeFile to make library fully platform-agnostic ([d8a0e11](https://github.com/sebastian-software/xlsx-format/commit/d8a0e119fbf2f87766d9ced37356a080a109936f))


### Documentation

* add browser-supported badge to README ([9e1cd48](https://github.com/sebastian-software/xlsx-format/commit/9e1cd48053668f86dd78ab269469577c33b575e5))

## [1.0.2](https://github.com/sebastian-software/xlsx-format/compare/xlsx-format-v1.0.1...xlsx-format-v1.0.2) (2026-02-12)


### Bug Fixes

* ardo version issues ([0ea9cf9](https://github.com/sebastian-software/xlsx-format/commit/0ea9cf91a9939206491ad8701697b69231310e06))
* read project metadata from root package.json for footer ([a85e081](https://github.com/sebastian-software/xlsx-format/commit/a85e08117f943d96049c835544bc0db1274afd24))
* restore Shiki theme override until ardo vite plugin default is fixed ([a1fa1a5](https://github.com/sebastian-software/xlsx-format/commit/a1fa1a5b45101dba09b552a5a09eb9cede818781))


### Documentation

* add API documentation browser with GitHub Pages deployment ([4d7a033](https://github.com/sebastian-software/xlsx-format/commit/4d7a0332b9c6e00a1e0f93e9cf926f7cfbfbbb77))
* add version display, build timestamp, and sponsor link to footer ([e19b469](https://github.com/sebastian-software/xlsx-format/commit/e19b46921a7eb73295810d1044e4a7f772ed9c0f))
* overhaul documentation site with real content and Blue Steel theme ([e83ec4e](https://github.com/sebastian-software/xlsx-format/commit/e83ec4e9b2f1cc71a34f279b1286d28b341d5ac3))

## [1.0.1](https://github.com/sebastian-software/xlsx-format/compare/xlsx-format-v1.0.0...xlsx-format-v1.0.1) (2026-02-11)


### Bug Fixes

* exclude __fixtures__ from coverage measurement ([a0efe16](https://github.com/sebastian-software/xlsx-format/commit/a0efe1647832a40bce7e612e3d85e96d063a6ce8))


### Refactoring

* co-locate tests next to source files and fix lint/formatting ([84d96d9](https://github.com/sebastian-software/xlsx-format/commit/84d96d9c0835a2f33eafd2a5f46d87b5c422f276))


### Documentation

* add feature support table to README ([d26a69b](https://github.com/sebastian-software/xlsx-format/commit/d26a69baa192776ee9f3ea1e14a14c5528791cd9))
* add test coverage row to comparison table ([7a1054f](https://github.com/sebastian-software/xlsx-format/commit/7a1054f923173ca483d12865e17467e18deefc30))
* rewrite README with sharper copy and better structure ([5fcd852](https://github.com/sebastian-software/xlsx-format/commit/5fcd852d7d3746902ff4b7a69c5ab15a9cc5c454))
