# AGENTS.md

## Build & Run

```bash
dotnet build          # .NET 10.0 SDK (pinned in global.json)
dotnet run -- --tables path/to/__tables__.xlsx   # pass tables config
```

The tool is a single-project console app (`ExcelTool.csproj` → `ExcelTool.sln`). No tests, no CI.

## Excel Sheet Convention

| Row (0-indexed) | Purpose |
|---|---|
| 0 | Field names |
| 1 | Field types (case-insensitive) |
| 2 | Field descriptions (→ XML doc comments) |
| 5+ | Data rows |

- Sheets whose name starts with `#` are skipped entirely.
- The sheet named `__enums__` is parsed specially — not as a data table but as enum definitions.
- Empty data rows are skipped automatically.

## The `__tables__.xlsx` Master Config

Column layout (row 0-3 are ignored; data starts at row 4, 0-indexed):

| Col | Field | Example |
|---|---|---|
| A | Input .xlsx path | `AudioConfigs.xlsx` |
| B | Namespace | `OCES.Audio` |
| C | Output code dir | `../hola_unity/Assets/Src/Config/` |
| D | Output data dir | `../hola_unity/Assets/Resources/Config/` |

Paths in the config are relative to the directory containing `__tables__.xlsx` unless rooted.

## Type System

All supported types are registered in `TypeRegistry.cs`. To add a new type, add a `Register(...)` call in `RegisterAll()`. The registry drives both binary serialization and C# code generation — one change fixes both.

Supported types: `int`, `uint`, `short`, `ushort`, `sbyte`, `byte`, `float`, `double`, `long`, `bool`, `string`, `vector` (→ `List<float>`, 3 components), `vectorlist` (→ `List<List<float>>`), `xxxlist` (e.g., `intlist`), `list<T>` generic syntax.

Any type **not** in the registry is treated as an enum and serialized as a single `byte`.

## Output

Per parsed sheet, three artifacts are generated:
1. `{SheetName}.cs` — data model class + `{SheetName}Config` container (both `partial`, implementing `IBinarySerializable`)
2. `AudioConsts.cs` — enums/Cues/NameDictionaries (only if `AudioObject` sheet or `__enums__` sheet exists)
3. `{SheetName}.bytes` — binary serialized data for Unity `Resources.Load<TextAsset>`

The first field of every data sheet is assumed to be `Id` (used as dictionary key in `*Config.QueryById()`).

## Key Files

| File | Role |
|---|---|
| `Program.cs` | CLI entrypoint (`--tables` option) |
| `TablesConfig.cs` | Parses `__tables__.xlsx` into `TableEntry` list |
| `ExcelHelper.cs` | Parses individual .xlsx → `ParsedSheet` / `ParsedEnum` |
| `TypeRegistry.cs` | Central type registry (binary write + codegen) |
| `Parser/GenModels.cs` | Generates `{SheetName}.cs` |
| `Parser/GenAudioConsts.cs` | Generates `AudioConsts.cs` |
| `Parser/TableExcelExportBytes.cs` | Generates `.bytes` files |
| `FileManager.cs` | File I/O utilities |
