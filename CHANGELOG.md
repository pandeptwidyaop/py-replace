## [1.0.4](https://github.com/pandeptwidyaop/py-replace/compare/v1.0.3...v1.0.4) (2025-11-08)

### 🐛 Bug Fixes

* **macos:** fix template path resolution for macOS app bundle ([dd6a983](https://github.com/pandeptwidyaop/py-replace/commit/dd6a98304bb15ab0db58998ebd80c975ed09c5dd))

## [1.0.3](https://github.com/pandeptwidyaop/py-replace/compare/v1.0.2...v1.0.3) (2025-11-08)

### 🐛 Bug Fixes

* **build:** add runtime hook to fix python-docx template path resolution ([dd675a7](https://github.com/pandeptwidyaop/py-replace/commit/dd675a73f17831e3cd8f1c1285ad8bb00beeb091))
* **build:** collect all docx submodules for proper template resolution ([902c1cb](https://github.com/pandeptwidyaop/py-replace/commit/902c1cba3bff233b0e5d94ed2d0f95d810d3b576))

## [1.0.2](https://github.com/pandeptwidyaop/py-replace/compare/v1.0.1...v1.0.2) (2025-11-08)

### 🐛 Bug Fixes

* **build:** switch to onedir mode to fix template path resolution ([cbc542c](https://github.com/pandeptwidyaop/py-replace/commit/cbc542c82d1cf84379d1e2c0cc4fbeac46ac862c))
* **build:** use collect_data_files for python-docx templates ([b0807f0](https://github.com/pandeptwidyaop/py-replace/commit/b0807f070f89ba5d09af5393b9fc0fe1f9bb8554))
* **ci:** update packaging for onedir build mode ([f346b04](https://github.com/pandeptwidyaop/py-replace/commit/f346b04458d02f7f6f5196e08f3dc4cc2d831340))

## [1.0.1](https://github.com/pandeptwidyaop/py-replace/compare/v1.0.0...v1.0.1) (2025-11-08)

### 🐛 Bug Fixes

* include python-docx template path in spec file ([a774796](https://github.com/pandeptwidyaop/py-replace/commit/a774796e2d98e4161a9285feb8108587bcb16809))

## 1.0.0 (2025-11-08)

### ⚠ BREAKING CHANGES

* Workflow now triggers on push to main branch instead of version tags. Use conventional commits for automated versioning.

🤖 Generated with [Claude Code](https://claude.com/claude-code)

Co-Authored-By: Claude <noreply@anthropic.com>

### ✨ Features

* implement semantic release workflow with automated versioning ([f044803](https://github.com/pandeptwidyaop/py-replace/commit/f044803473da4f516792e869c4925ba35adfce35))

### 🐛 Bug Fixes

* **ci:** add missing @semantic-release/github plugin and improve error handling ([1b45e99](https://github.com/pandeptwidyaop/py-replace/commit/1b45e993ef8c0b7935c83b950c3e69d1a0dc8be3))
* **ci:** add workflow permissions for semantic-release to push ([3535049](https://github.com/pandeptwidyaop/py-replace/commit/353504905fd9bbb06e244b306f767a21fc8e3a77))
* **ci:** improve version detection pattern for semantic-release v24 ([ecd5357](https://github.com/pandeptwidyaop/py-replace/commit/ecd535743e7668235b0c63a21205465b1659766b))
* missing ([9325abb](https://github.com/pandeptwidyaop/py-replace/commit/9325abbba6933b3d6b86f2ee267c3beb5b6aebd4))
* update repository URL in configuration and documentation ([62fef40](https://github.com/pandeptwidyaop/py-replace/commit/62fef40a098840ae4f8aa083bfbabf7451af5c08))
