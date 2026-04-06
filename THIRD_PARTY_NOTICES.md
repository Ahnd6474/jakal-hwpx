# Third-Party Notices

The top-level [LICENSE](./LICENSE) applies to original source code in this repository unless a file, directory, or notice says otherwise.

## HWP and HWPX Compatibility Notice

This project is an independent compatibility tool for working with `HWP` and `HWPX` documents.

`Hancom`, `Hangul`, `HWP`, and `HWPX` are referenced only to describe file-format compatibility. No affiliation with, sponsorship by, or endorsement from the respective rights holders is claimed.

## Sample and Fixture Documents

Sample or fixture documents in this repository, including files under `examples/`, committed generated outputs, and any repository-root `*.hwp`, `*.hwpx`, or `*.pdf` files, may carry separate copyright, trademark, privacy, publicity, or template rights.

Unless a file explicitly says otherwise, treat those documents as test or reference assets only. They are not automatically relicensed under the project's MIT license.

## Bundled Third-Party Tools

- `tools/jdk-21.0.10+7/` and `tools/OpenJDK21U-jdk_x64_windows_hotspot_21.0.10_7.zip` are redistributed OpenJDK or Eclipse Temurin artifacts and remain under their own upstream license and notice files shipped with those assets.
- `tools/lib/hwplib-1.1.10.jar` contains Maven metadata declaring Apache-2.0 for `kr.dogfoot:hwplib:1.1.10`.
- `tools/lib/hwpxlib-1.0.8.jar` contains Maven metadata declaring Apache-2.0 for `kr.dogfoot:hwpxlib:1.0.8`.
- `tools/hwp-batch-converter.jar` is a bundled Java converter artifact used for maintainer workflows. It should be treated as a separate toolchain artifact together with its bundled or referenced upstream components, not as blanket-MIT project source.
- `tools/java-src/local/jakaldocs/HwpBatchConverterMain.java` is project-authored integration code and is covered by the MIT license unless noted otherwise.

## Redistribution Guidance

If you redistribute this repository or a derivative:

- keep the top-level `LICENSE` with project-authored code
- keep upstream license and notice files for bundled third-party tools
- confirm that you have the right to redistribute included sample documents
- avoid using third-party product or trademark names in a way that suggests endorsement
