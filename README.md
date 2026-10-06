# Excel to PostgreSQL Bulk Loader (2 TB Benchmark)

A Python program that loads a very large Excel dataset, made of many workbooks with
many worksheets each, into a single PostgreSQL table and measures the actual
execution time of the load.

## 1. Project Overview

The task:

> Write a program in Python to load 2 TB data into a table. Please try with 2 TB data
> and give me the execution time.
>
> In Excel, for 1 sheet you have a row limit, but you can add N number of sheets in
> Excel to have 2 TB data and try to load.

`task_one.py` is the program for this task. It:

- finds every Excel workbook (`.xlsx`) under an input directory, including subdirectories,
- treats every worksheet as one partition of the dataset,
- reads each worksheet incrementally, validates it against the target table and
  bulk-loads it into PostgreSQL with `COPY`, batch by batch,
- measures the run (start time, end time, execution time, rows, bytes, throughput)
  and prints a benchmark summary,
- when the dataset is smaller than 2 TB, also prints a 2 TB estimate that is clearly
  labelled **Projected** and kept apart from the **Actual** measurements.

`generate_test_data.py` creates a synthetic multi-sheet Excel dataset of any size for
the benchmark.

## 2. Problem

An Excel worksheet holds at most **1,048,576 rows** (one of them is the header) and
16,384 columns. A single sheet therefore cannot hold a 2 TB dataset; the data has to
be split across many worksheets and, in practice, many workbooks.

For scale: with the 7-column layout produced by `generate_test_data.py`, a full
worksheet (1,048,575 data rows) is roughly 50 MB of `.xlsx`, so 2 TB of Excel is in
the order of 40,000 full worksheets, for example 4,000 workbooks with 10 sheets each.
The generator prints the exact size of the dataset it creates.

This project uses that structure directly: **every worksheet is a partition** of the
overall dataset, and the loader processes the partitions one after another (and in
parallel across workbooks), so the total size is limited only by disk space and time.

An `.xlsx` file is a ZIP archive of XML documents. Reading it means decompressing and
parsing XML, which is CPU-bound. At multi-terabyte scale **reading Excel is the main
bottleneck**, not PostgreSQL. Section 12 covers what this means for a 2 TB run.

## 3. Technical Approach

**Sheets as partitions.** The input directory is searched for workbooks (`*.xlsx` by
default). Each workbook is processed by one worker process, which walks through its
worksheets in order. The unit of loading, commit and resume is one worksheet.

**Streaming, chunked reading.** Two Excel readers are available through `--engine`:

| Engine | How it reads | Memory per worker | Speed |
| ------ | ------------ | ----------------- | ----- |
| `openpyxl` (default) | Read-only streaming mode: rows are parsed from the sheet XML one at a time | Flat, tens of MB, independent of sheet and file size | Pure Python, slow |
| `calamine` | Rust-based reader (`python-calamine`) that loads one worksheet at a time | Grows with the largest sheet; capped by Excel's row limit (several hundred MB for a full sheet with a handful of columns) | Several times faster end to end |

Neither engine loads the whole workbook, let alone the whole dataset, into memory.
`pandas.read_excel` is not used.

**Bounded memory.** Rows flow through a pipeline of generators: reader, validation,
batching, `COPY`. A batch is never materialised as a list; each row is written to the
`COPY` stream as soon as it is validated, so memory does not grow with the batch size,
the sheet size or the dataset size (apart from the per-sheet memory of the `calamine`
engine described above).

**Batch processing.** Each worksheet is sent to PostgreSQL in batches of
`--batch-size` rows (default 50,000). Each batch is one `COPY` command.

**PostgreSQL bulk loading.** Data is loaded with `COPY ... FROM STDIN` through
psycopg 3, never with row-by-row `INSERT`s. Sessions run with
`synchronous_commit=off`, which is safe here because a sheet's data and its checkpoint
are committed together (see below). A missing target table can be created as
`UNLOGGED` with `--unlogged` for extra speed. The program never creates indexes on the
target table; for the fastest load, add indexes after the data is in.

**Schema validation.** The first non-empty row of each sheet is its header. Header
names are matched to the target table's columns case-insensitively, ignoring spaces
and punctuation (`Order ID` matches `order_id`). A sheet fails validation, and is not
loaded, if a header is blank, duplicated or not a column of the table, or if a
`NOT NULL` column without a default is missing. If the target table does not exist, it
is created from the first sheet's header, with column types inferred from up to 1,000
sample rows (`bigint`, `double precision`, `boolean`, `timestamp`, `time`, `interval`
or `text`). Create the table yourself when you need exact types.

**Data validation.** Every cell is checked against the type of its target column:
integers (including range checks for `smallint`, `integer` and `bigint`), numeric and
floating point, boolean, date, timestamp, time, and text (including the length limit
of `varchar(n)` and `char(n)`). Values for other column types are passed to PostgreSQL
unchanged. Blank cells become `NULL`. A row that fails validation is **rejected**: it
is skipped, counted in the summary, and the first five rejections per sheet are logged
with their row number and reason. The rest of the sheet still loads.

**Transactions, checkpoints and resume.** All batches of a worksheet and a row in the
`excel_load_checkpoint` table (created next to the target table) are committed in one
transaction. A worksheet is therefore either completely loaded and recorded, or not
loaded at all, never half loaded. If a run stops (crash, reboot, Ctrl+C), running the
same command again skips the worksheets that are already loaded and continues with the
rest. `--truncate` empties the target table and its checkpoints for a clean benchmark.

**Error handling and retries.** A lost connection or other server-side failure
(PostgreSQL `OperationalError`) rolls the worksheet back and retries it up to
`--retries` times with exponential backoff (2 s, 4 s, 8 s, and so on). Errors that a
retry cannot fix (a schema mismatch, a data error reported by PostgreSQL, an
unreadable worksheet) fail that worksheet only, and the run continues with the next
one. An unreadable workbook (corrupt file, not a real `.xlsx`) is reported as a failed
file. Failures are listed in the log, counted in the summary, and make the program
exit with code 1. Connections and files are always closed.

**Parallelism.** `--workers` worker processes load different workbooks at the same
time; separate processes, rather than threads, are needed because Excel parsing is
CPU-bound Python code. The default is half the logical CPUs, because each worker also
keeps a PostgreSQL backend busy. More workers are not always faster: once CPU, disk or
PostgreSQL is saturated, extra workers only compete with each other. Measure with a
few values on the target machine.

**Benchmark measurement.** The run is timed with a monotonic high-resolution clock
(`time.perf_counter`). Start and end times are taken from the wall clock and shown
with their UTC offset. All counts are collected while the data is being loaded;
nothing in the summary is hard-coded or estimated except the clearly labelled
projection.

## 4. Architecture / Workflow

```text
Excel Files
    ↓
Discover Files            recursive search for *.xlsx, Excel lock files (~$) skipped
    ↓
Discover Sheets           one worker process per workbook, worksheets in order
    ↓
Read Sheet Incrementally  openpyxl read-only stream, or calamine one sheet at a time
    ↓
Validate Data             header against table columns, cell types, required values
    ↓
Create Batch              --batch-size rows, streamed, never held in memory as a list
    ↓
PostgreSQL Bulk Load      COPY ... FROM STDIN
    ↓
Record Progress           checkpoint row in the same transaction, progress log line
    ↓
Next Batch
    ↓
Next Sheet
    ↓
Benchmark Summary         start, end, execution time, throughput, 2 TB projection
```

## 5. 2 TB Benchmark

### Status

**The real 2 TB benchmark result is pending.** No result has been invented or
estimated in this README. The table below is filled in only with values measured by
`task_one.py`.

### Results

| Dataset            | Size | Rows | Execution Time | Throughput | Result                 |
| ------------------ | ---: | ---: | -------------: | ---------: | ---------------------- |
| Test dataset       |  TBD |  TBD |            TBD |        TBD | Actual                 |
| 2 TB Excel dataset | 2 TB |  TBD |            TBD |        TBD | Actual / Not completed |
| 2 TB projection    | 2 TB |  TBD |            TBD |        TBD | Projected              |

- Machine used: TBD (CPU, RAM, disk type, PostgreSQL version and location)
- Settings used: TBD (`--engine`, `--workers`, `--batch-size`)
- If the full 2 TB run was not completed, the reason: TBD

Fill the table from the benchmark summary or from the JSON written by `--report-file`.
The third row is only used when the full 2 TB run could not be completed; take its
values from the "Projected" lines of the summary of the largest run you completed.

### Running the 2 TB test

1. **Check disk space.** You need space for the 2 TB of `.xlsx` files **and** for
   PostgreSQL. Excel files are compressed, so the loaded table is larger than the
   input: in local tests with the generated data it was roughly twice the input size.
   Plan for about 2 TB of input plus 4 TB or more for the database and its WAL, ideally
   on separate disks. The summary reports the table growth and projects it to 2 TB.
2. **Prepare the dataset.** Use real Excel data, or generate it (see
   [Generating test data](#generating-test-data)), for example into `D:\2TB_DATA`.
3. **Calibrate on a smaller dataset first** (for example 10 to 50 GB) and try a few
   `--workers` values and both engines to find the fastest settings for the machine.
4. **Run the full load** with a clean table and a report file:

   ```text
   python task_one.py --input-dir "D:\2TB_DATA" --table my_target_table --workers 4 --engine calamine --truncate --report-file benchmark_report.json --log-file benchmark.log
   ```

   Use `--engine openpyxl` (the default) instead if the machine does not have enough
   RAM for `workers × the largest sheet` (see section 3).
5. **Record the result.** Copy the actual values from the summary into the table above.

### Actual versus projected

The summary always reports what was measured. If the processed data is 2 TB or more,
it states `RESULT: ACTUAL 2 TB RUN`: the execution time shown is the real 2 TB time.

If less than 2 TB was processed, because of disk space, RAM, CPU, I/O, time or
PostgreSQL limits, the summary reports:

- the **actual** dataset size tested, rows processed, execution time and throughput,
- a **projected** 2 TB load time (2 TB ÷ measured throughput), row count and
  PostgreSQL table size, each line marked `Projected`.

A projection assumes the measured throughput stays constant at 2 TB. It is an
estimate, not a benchmark result, and must always be reported as "Projected" together
with the reason the full run was not done.

## 6. Execution Time

The program measures the real duration of the run:

```text
End Time - Start Time = Actual Execution Time
```

- Start Time is taken when the run begins (before files are discovered), End Time
  when the last workbook has been processed. Both are printed with their UTC offset.
- The duration is measured with a monotonic clock, so it stays correct even if the
  system clock changes during a multi-day run (time sync, daylight saving).
- Throughput = Data Processed ÷ Execution Time. Rows/Second = Rows Processed ÷
  Execution Time.

No execution time is written in this README; it is produced by running the program.

## 7. Setup

Requirements: Python 3.10 or newer (CI tests 3.10 and 3.11) and PostgreSQL 12 or
newer.

On Windows (Command Prompt or PowerShell):

```text
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

If PowerShell refuses to run the activation script, allow local scripts once with
`Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`.

On Linux or macOS, activate with `source .venv/bin/activate` instead.

## 8. Configuration

Connection settings come from command-line options or the standard PostgreSQL
environment variables. Nothing is hard-coded and no password is stored in the code.

| Setting  | Option     | Environment variable | Default     |
| -------- | ---------- | -------------------- | ----------- |
| Host     | `--host`   | `PGHOST`             | `localhost` |
| Port     | `--port`   | `PGPORT`             | `5432`      |
| Database | `--dbname` | `PGDATABASE`         | `postgres`  |
| User     | `--user`   | `PGUSER`             | `postgres`  |
| Password | (none)     | `PGPASSWORD`         | (none)      |

The password is deliberately not accepted as a command-line option, because command
lines end up in shell history and process lists. Use one of these instead:

- set `PGPASSWORD` for the current terminal session only:

  ```text
  PowerShell:      $env:PGPASSWORD = "your-password"
  Command Prompt:  set PGPASSWORD=your-password
  ```

- or store it in a PostgreSQL password file, which is read automatically:
  `%APPDATA%\postgresql\pgpass.conf` on Windows, `~/.pgpass` elsewhere, one line per
  server in the form `host:port:database:user:password`.

Never commit a password or a `.env` file; `.gitignore` already excludes `.env`.

The database user needs permission to create tables (when the target table or the
checkpoint table does not exist yet) and to insert into and truncate the target table.

## 9. Running the Program

Load every workbook under `D:\2TB_DATA` into `my_target_table`:

```text
python task_one.py --input-dir "D:\2TB_DATA" --table my_target_table --workers 4
```

A clean benchmark run with the faster engine, a JSON report and a log file:

```text
python task_one.py --input-dir "D:\2TB_DATA" --table my_target_table --workers 4 --engine calamine --truncate --report-file benchmark_report.json --log-file benchmark.log
```

Resume an interrupted run by running the same command again without `--truncate`.

| Option | Description | Default |
| ------ | ----------- | ------- |
| `--input-dir` | Directory with the Excel workbooks; subdirectories are included | required |
| `--table` | Target table as `table` or `schema.table`; created from the data if missing | required |
| `--pattern` | Workbook file name pattern, for example `"*.xls[xm]"` to include `.xlsm` | `*.xlsx` |
| `--workers` | Worker processes, each loading one workbook at a time (at most 61 on Windows) | half the logical CPUs |
| `--batch-size` | Rows per `COPY` batch | `50000` |
| `--engine` | Excel reader: `openpyxl` (streaming) or `calamine` (faster) | `openpyxl` |
| `--retries` | Retries per sheet after a lost connection or server error | `3` |
| `--truncate` | Empty the target table and its checkpoints before loading | off |
| `--unlogged` | Create a missing target table as `UNLOGGED` (faster, not crash-safe) | off |
| `--report-file` | Also write the benchmark summary as JSON | none |
| `--log-file` | Also write the log to a file | none |
| `--verbose` | Log every batch | off |
| `--host`, `--port`, `--dbname`, `--user` | PostgreSQL connection, see section 8 | see section 8 |

Run `python task_one.py --help` for the same list.

### Benchmark summary

At the end of every run the program prints a summary in this format (values are
placeholders here; real values are measured during the run):

```text
================================================================
EXCEL TO POSTGRESQL BENCHMARK
================================================================
Start Time        : <measured>
End Time          : <measured>
Execution Time    : <measured>
Files Processed   : <measured>
Sheets Processed  : <measured>
Rows Processed    : <measured>
Data Processed    : <measured>
Throughput        : <measured>
Rows/Second       : <measured>
Successful Batches: <measured>
Failed Batches    : <measured>
----------------------------------------------------------------
Rows Rejected     : <measured>
Sheets Failed     : <measured>
Sheets Skipped    : <measured> (loaded by an earlier run)
Files Failed      : <measured>
Target Table      : <table>
Table Growth      : <measured>
Engine            : <setting>
Workers           : <setting>
Batch Size        : <setting>
================================================================
RESULT: ACTUAL for the dataset above. 2 TB figures below are PROJECTED.
Tested Dataset    : <measured> - Actual
2 TB Load Time    : <calculated> - Projected
2 TB Row Count    : <calculated> - Projected
2 TB Table Size   : <calculated> in PostgreSQL - Projected
Projections assume the measured throughput stays constant.
================================================================
```

| Field | Meaning |
| ----- | ------- |
| Files Processed | Workbooks read in this run |
| Sheets Processed | Worksheets loaded (or found empty) in this run |
| Rows Processed | Rows committed to PostgreSQL in this run |
| Data Processed | Size on disk of the workbooks read in this run |
| Throughput | Data Processed ÷ Execution Time |
| Successful Batches | `COPY` batches committed |
| Failed Batches | Batches rolled back because their worksheet failed |
| Rows Rejected | Rows skipped because they failed validation |
| Sheets Skipped | Worksheets already loaded by an earlier run (resume) |
| Table Growth | Growth of the target table on disk (data and indexes) during the run |

Sizes use decimal units: 1 MB = 10^6 bytes and 1 TB = 10^12 bytes. Windows Explorer
uses binary units, so it shows 2 TB as about 1.82 TB.

Exit codes: `0` when everything loaded, `1` when any workbook or worksheet failed (or
there was nothing to load), `130` when the run was interrupted with Ctrl+C.

### Generating test data

`generate_test_data.py` writes unique workbooks with the requested number of sheets
and rows (order data: id, customer, region, product, quantity, price, date). Writing
terabytes of unique rows with openpyxl would take weeks, so `--target-size` adds
byte-for-byte copies of the generated workbooks until the dataset reaches the requested
size. Copies repeat the `order_id` values of the workbook they copy; that does not
change the work needed to read and load them. Before copying, the generator checks
that the disk has enough free space for the target size and stops if it does not.
While a workbook is being written, openpyxl keeps each sheet as an uncompressed
temporary file in the system temp directory, so generating workbooks of full sheets
needs several GB of temporary space per workbook in progress.

A quick 1 GB test dataset:

```text
python generate_test_data.py --output-dir D:\excel_test --files 2 --sheets-per-file 5 --rows-per-sheet 200000 --target-size 1GB
```

A 2 TB dataset of full worksheets:

```text
python generate_test_data.py --output-dir "D:\2TB_DATA" --files 4 --sheets-per-file 10 --rows-per-sheet 1048575 --target-size 2TB
```

## 10. Running Tests

```text
pytest
```

The unit tests need no database and no large files: they build small workbooks in
temporary directories and use a fake PostgreSQL connection. They cover file discovery,
sheet processing with both engines, header and type validation, batching, transactions
and retries, error handling, empty and invalid input, time measurement, throughput,
the summary and projection, the command line and the data generator.

`test_task_one.py` also contains an end-to-end test that loads real workbooks into a
real PostgreSQL database with two worker processes, then checks the row counts,
inferred column types, resume and `--truncate`. It is skipped unless
`RUN_POSTGRES_TESTS=1` is set; it connects using the `PG*` environment variables and
works in a temporary schema that it drops afterwards. The GitHub Actions workflow runs
it against a PostgreSQL service container.

## 11. Code Quality

```text
flake8 .
black --check .
```

GitHub Actions (`.github/workflows/ci.yml`) runs `pytest`, `flake8 .` and
`black --check .` on Python 3.10 and 3.11 for every push and pull request.

## 12. Limitations

- **Excel's per-sheet row limit.** A worksheet holds at most 1,048,576 rows, so 2 TB of
  Excel means tens of thousands of worksheets spread over thousands of workbooks.
- **Very large Excel files.** Excel itself cannot open multi-gigabyte workbooks;
  this program can, because it never needs to load a whole workbook. Workbooks saved by
  Excel keep their text in a shared-strings table, which both engines load into memory
  for each open workbook. Very large workbooks with many unique strings therefore need
  more RAM per worker; many medium-sized workbooks (up to a few hundred MB each) work
  best. Formulas are loaded as their last calculated values. The header must be the
  first non-empty row of each worksheet. Legacy `.xls` files must be converted to
  `.xlsx` first.
- **Disk space.** A full run needs room for the 2 TB input and for the loaded table,
  which is larger than the compressed `.xlsx` input, plus PostgreSQL's write-ahead log.
- **RAM.** The `openpyxl` engine uses little memory per worker. The `calamine` engine
  holds one worksheet per worker in memory, so its RAM need grows with
  `--workers` × the largest sheet.
- **CPU.** Decompressing and parsing Excel XML is the bottleneck. It scales with the
  number of cores through `--workers`, and PostgreSQL also needs CPU for `COPY`; a
  2 TB load on a laptop-class CPU takes days to weeks.
- **Disk I/O.** Reading 2 TB of input and writing several TB to PostgreSQL at the same
  time needs fast storage; input and database on different disks help.
- **PostgreSQL performance.** Indexes, constraints, triggers and foreign keys on the
  target table slow `COPY` down considerably. Load into a table without them and
  create indexes afterwards. `--unlogged` avoids WAL writes but the table is emptied
  after a crash and is not replicated; run `ALTER TABLE ... SET LOGGED` after loading
  if it must be durable. Settings such as `max_wal_size` and `checkpoint_timeout` are
  server-level and should be tuned by the DBA for a load of this size.
- **Network throughput.** When PostgreSQL runs on another machine, the `COPY` data
  crosses the network; the link must be fast enough not to become the bottleneck.
- **Actual versus projected.** Only a run that really processes 2 TB gives the actual
  2 TB execution time. A projection from a smaller run assumes constant throughput;
  in reality it can drop as the table grows, disk caches stop helping and
  checkpoints and autovacuum add work, so the projection is a lower-bound estimate.
- **Resume granularity.** A worksheet interrupted in the middle is loaded again from
  its first row on the next run. Throughput figures are meaningful for complete runs;
  start benchmark runs with `--truncate`.

## Project Structure

```text
.
├── task_one.py             Loader and benchmark (main program)
├── generate_test_data.py   Synthetic multi-sheet Excel dataset generator
├── test_task_one.py        pytest suite
├── requirements.txt        Pinned runtime and development dependencies
├── .flake8                 flake8 settings compatible with black
├── .gitignore
├── LICENSE                 MIT
└── .github/
    └── workflows/
        └── ci.yml          GitHub Actions: pytest, flake8, black on Python 3.10 and 3.11
```

## License

Released under the [MIT License](LICENSE).
