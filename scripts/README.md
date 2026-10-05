# Repository scripts

Automation is grouped by responsibility instead of keeping every entry point at
the `scripts/` root:

| Directory     | Purpose                                                                        |
| ------------- | ------------------------------------------------------------------------------ |
| `corpus/`     | Third-party corpus setup, round-trip gates, baseline, and semantic comparison. |
| `coverage/`   | OOXML, ODF, and legacy capability coverage plus their registries.              |
| `gates/`      | Package, output, container, cross-format, and documentation parity checks.     |
| `models/`     | Generate and verify the XSD-derived container content models.                  |
| `schema/`     | Generate JSON schemas and validate options corpora.                            |
| `validation/` | Generate demos and validate OOXML/ODF XML against golden schemas.              |
| `dev/`        | Developer conveniences that are not CI gates.                                  |
| `fixtures/`   | Compressed cross-format fixtures.                                              |

Prefer the named package scripts in `package.json`; they are the stable public
entry points. Add a new script under the matching directory and document it here.
