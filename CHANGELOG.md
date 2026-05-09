# xlsx-kit

## 0.2.0

### Minor Changes

- [`b36ca45`](https://github.com/baseballyama/xlsx-kit/commit/b36ca453b08c91981baac42b3b5bc4aeeeef6ec0) Thanks [@baseballyama](https://github.com/baseballyama)! - Hardening and docs release.

  - Add a 3-tier ECMA-376 conformance validator and broaden conformance coverage to the writer surface, real-world fixtures, and fast-check property tests.
  - Add `knip` to CI to keep the public export surface tight; prune unused exports flagged by it.
  - Refresh the docs site: redesigned landing and docs UI with a new typography system, new logo and favicons, and new "Why xlsx-kit" / comparison / motivation sections in the README.
  - Tighten release / dependency automation: pin dependencies, drop EOL Node 18/20 from the test matrix and add Node 26, bump the project Node engine to 22.22.2.
