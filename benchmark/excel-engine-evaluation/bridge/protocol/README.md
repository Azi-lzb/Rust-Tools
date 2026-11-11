# Worker protocol v1 semantics

`excel-task-v1.schema.json` describes the common path-based request envelope. The implemented Rust→Aspose .NET and Rust→openpyxl prototypes are deliberately narrower and do not claim support for every schema operation. A worker must advertise its capabilities and reject unknown operations with `unsupported_operation`; it must never silently skip a requested operation.

## Ordering and coordinates

- `operations` execute in array order. `calculate=true` means calculate after all edits and before save. `save_format=preserve` retains the source format; explicit `xls`/`xlsx` requests a conversion and must be rejected if the engine cannot perform it.
- Rows and columns are zero-based in the protocol. A C++/.NET/Python adapter translates at its boundary to the SDK convention. A request with a negative or overflowing coordinate fails before modifying a candidate.
- `append_rows.start_row` is the zero-based first destination row. If omitted in a future protocol revision, the engine must use the first row after the used range. Rows are rectangular logically; missing trailing elements are treated as `null` (clear/no value), not as a request to preserve old content.

## Values and strings

- Strings are UTF-8 JSON strings. `""` is an explicit empty string; `null` clears the cell value; an omitted cell in a sparse `set_values` list means no operation. Styled blank cells are represented by their style operation, not by a value.
- JSON numbers use IEEE-754 binary64. Implementations must reject non-finite values and preserve round-trip precision available in the file format.
- Dates are explicit `{ "type": "date_serial", "value": n }` Excel serials. The workbook's 1900/1904 date system controls interpretation. Adapters must not silently reinterpret local time zones.
- Formula text is sent as a string beginning with `=`. It is written as a formula, not a literal. Calculation is a separate capability; `calculate=true` on an engine without it returns `unsupported_capability`.

## Safety and errors

- `input_path` and `output_path` must resolve to distinct files. The source is snapshotted before work; output is written to a sibling candidate, reopened and validated, then atomically renamed to the requested output. A worker never modifies the input in place.
- Each response carries `protocol_version`, `request_id`, `ok`, and either result metadata or a structured error `{code,message,engine,stage,retryable}`. Standard codes include `unsupported_operation`, `unsupported_capability`, `invalid_request`, `input_changed`, `load_failed`, `save_failed`, `validation_failed`, `cancelled`, and `timeout`.
- Timeouts are enforced by the Rust supervisor, which terminates the controlled process tree. Cancellation is best-effort and must leave the candidate unpublished. Diagnostics go to stderr; stdout is protocol-only for streaming workers.
- Worker capability declarations identify supported formats, operations, formula calculation, formula families, and maximum tested workbook size. `NOT_TESTED` is not the same as unsupported.

## Implemented prototype subsets

- Rust→LibXL C ABI: opaque workbook handle; open, bulk append rows, save, reopen and verify. `lastRow()` was adapted as a used-row count after the first indexing assertion failed. It does not implement the complete protocol.
- Rust→Aspose .NET: one resident self-contained worker, JSON-lines on stdin/stdout, input/output paths and one batched append request. It validates that the source and output paths differ, writes a sibling candidate and checks the last appended value.
- Rust→openpyxl: test-only JSON request/response files; one process per workbook task, one batch append per process, candidate save + atomic replace + readback. This is not a production worker modification.
