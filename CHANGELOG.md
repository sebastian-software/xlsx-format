# Changelog

## [0.21.0](https://github.com/sebastian-software/xlsx-format/compare/v0.20.3...v0.21.0) (2026-02-11)


### Features

* add CSV, TSV, and HTML as read/write targets ([f6df6ab](https://github.com/sebastian-software/xlsx-format/commit/f6df6abdc5bbd96da111e47713f5d5ad70e2f9b6))
* Add Sheet Protection for XLS (BIFF8) ([#3202](https://github.com/sebastian-software/xlsx-format/issues/3202)) ([6c0f950](https://github.com/sebastian-software/xlsx-format/commit/6c0f950f83a2bb96e9dccb79c29ba47fb5a573b3)), closes [#3201](https://github.com/sebastian-software/xlsx-format/issues/3201)


### Bug Fixes

* Add DenseSheetData type ([#3195](https://github.com/sebastian-software/xlsx-format/issues/3195)) ([6912bdb](https://github.com/sebastian-software/xlsx-format/commit/6912bdb2d449525d65244d461e7afe2670d57a75))
* infinite loop due to hidden row in XLSX.stream.to_json ([#3178](https://github.com/sebastian-software/xlsx-format/issues/3178)) ([5550b90](https://github.com/sebastian-software/xlsx-format/commit/5550b907041cb7a16e449b5bfeb371668be34d00))
* missing break condition in make_json_row ([9c3853b](https://github.com/sebastian-software/xlsx-format/commit/9c3853ba253cbf7b3fe9884667307dd66585f1f7))
