# Third-party workbooks

The benchmark and the tests use real workbooks written by other people. They are
not stored in this repository. `fetch_corpus.py` downloads each one from its
upstream source at a pinned commit and checks its sha256 before writing it:

    python research/xlsx-bench/fetch_corpus.py           # download what is missing
    python research/xlsx-bench/fetch_corpus.py --check   # verify only

Tests that need a workbook are skipped until you run it.

Both manifests carry the full URL, size and licence of every file:
`research/xlsx-bench/corpus.json` and `research/safexl-moat/fixtures/fixtures.json`.
The local file name can differ from the upstream name.

| File | Upstream repo | Upstream path | Commit | Licence | sha256 |
|---|---|---|---|---|---|
| `research/xlsx-bench/corpus/simple-monthly-budget.xlsx` | apache/poi | `test-data/spreadsheet/simple-monthly-budget.xlsx` | `942d95d` | Apache-2.0 | `cae00c6894b95743017aeded4186f15a12c2d8382d71760029df490d58197db6` |
| `research/xlsx-bench/corpus/benefits-contribution-calculator.xlsm` | apache/poi | `test-data/spreadsheet/mv-calculator-final-2-20-2013.xlsm` | `942d95d` | Apache-2.0 | `7e7121abc05437469fac86612ee8f3e3118a0ecd5f93f95fe984f4a68d01d73d` |
| `research/xlsx-bench/corpus/structured-refs-lookup-tool.xlsx` | apache/poi | `test-data/spreadsheet/StructuredRefs-lots-with-lookups.xlsx` | `942d95d` | Apache-2.0 | `09af82b9abae2ffcff04777b69cbb23dfd15df788426153bcfa1e9359c3552be` |
| `research/xlsx-bench/corpus/pivot-chart-report.xlsx` | apache/poi | `test-data/spreadsheet/WithChartSheet.xlsx` | `942d95d` | Apache-2.0 | `792b9f21afdb6b8b1f82345864ee15d55c17a1b54d71c250e82c87304385d362` |
| `research/xlsx-bench/corpus/pension-hedging-stress-model.xlsx` | ClosedXML/ClosedXML | `ClosedXML.Tests/Resource/TryToLoad/TemplateWithTableSourcePivotTables.xlsx` | `09f2f21` | MIT | `c61200baa55082139a1c9e5ccc159c100b708cd2bc5d5d1c2157fc018d629eaf` |
| `research/xlsx-bench/corpus/sparkline-external-reference.xlsx` | ClosedXML/ClosedXML | `ClosedXML.Tests/Resource/Other/Sparklines/SourceDataFormulas/ExternalReference-input.xlsx` | `09f2f21` | MIT | `5b1101cf2e493dbac74ef4643bcfc64346db1c3f6d28c569acd637c58eebb5a2` |
| `research/xlsx-bench/corpus/kyc-intake-data-dictionary.xlsx` | apache/poi | `test-data/spreadsheet/CVLKRA-KYC_Download_File_Structure_V3.1.xlsx` | `942d95d` | Apache-2.0 | `a6a7bf557df8048108502973ecedc31d96f27d09ac166eedeae32fc96c13e8db` |
| `research/xlsx-bench/corpus/parking-study-report.xlsm` | apache/poi | `test-data/spreadsheet/57181.xlsm` | `942d95d` | Apache-2.0 | `2cf7459faa9d15c1deef4c410fc05142c4f57ebee622668e2526d7b3e34e08fa` |
| `research/safexl-moat/fixtures/DataValidations-49244.xlsx` | apache/poi | `test-data/spreadsheet/DataValidations-49244.xlsx` | `942d95d` | Apache-2.0 | `8066b12c878d74aa35b55b22dc9daba9a56e91332a35cf96c4fc0f6af6bd3bd1` |
| `research/safexl-moat/fixtures/ExcelPivotTableSample.xlsx` | apache/poi | `test-data/spreadsheet/ExcelPivotTableSample.xlsx` | `942d95d` | Apache-2.0 | `f0ef1d61c9f3d8b27b18a876b8e24676314c6450ed64a725a7bfaa603fdd67e3` |
| `research/safexl-moat/fixtures/NewStyleConditionalFormattings.xlsx` | apache/poi | `test-data/spreadsheet/NewStyleConditionalFormattings.xlsx` | `942d95d` | Apache-2.0 | `1e97af3b87636bd0a759c25439a21b928c2ff3b9a3c57ca49ae24cb1ccdb7e2f` |
| `research/safexl-moat/fixtures/SimpleMacro.xlsm` | apache/poi | `test-data/spreadsheet/SimpleMacro.xlsm` | `942d95d` | Apache-2.0 | `f76c986f4ebc25c2cc57c088b2511a1269f4bd61d6223a2ab58db351da348ba6` |
| `research/safexl-moat/fixtures/WithChart.xlsx` | apache/poi | `test-data/spreadsheet/WithChart.xlsx` | `942d95d` | Apache-2.0 | `51f27aabf7417f5aaa0542857b6bb56a92d4afbabdc5c63f084c06f92b67429d` |
| `research/safexl-moat/fixtures/WithDrawing.xlsx` | apache/poi | `test-data/spreadsheet/WithDrawing.xlsx` | `942d95d` | Apache-2.0 | `97b1ab359d6aecbe37864f30c85106e386f0a19bea8f290103927362a2701066` |

Pinned commits (full hashes):

- apache/poi `942d95d85b15d0dfdb3bc9ba1b4f273f277757c8`
- ClosedXML/ClosedXML `09f2f21a60e821ce7fb18ca3576b272e5dd6d921`

One workbook is ours and is not third-party: `research/safexl-moat/fixtures/formulas_definednames.xlsx` was generated with openpyxl for the tests, is MIT like the rest of the repository, and is the only workbook committed to it.

## Licences

- Apache POI (Apache-2.0): [LICENSE](https://github.com/apache/poi/blob/942d95d85b15d0dfdb3bc9ba1b4f273f277757c8/legal/LICENSE), [NOTICE](https://github.com/apache/poi/blob/942d95d85b15d0dfdb3bc9ba1b4f273f277757c8/legal/NOTICE)
- ClosedXML (MIT): [LICENSE](https://github.com/ClosedXML/ClosedXML/blob/09f2f21a60e821ce7fb18ca3576b272e5dd6d921/LICENSE)

Copies of these three files, as they stood at the pinned commits, are in `third_party_licenses/`.

The recordings and results in this repository quote cell contents and part names from these workbooks. Those excerpts stay under the workbooks' own licences.
