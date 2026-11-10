# UFT Automation Framework

A reusable, data-driven web automation framework for **OpenText UFT One / Functional Testing**, ready to clone into a new project. All URLs, credentials and settings live in a `.env` file; nothing is hard-coded to a machine path.

## Folder structure

```
UFT_Automation_Framework/
├── .env                  Local config with real URLs/credentials (git-ignored)
├── .env.example          Committed template - copy to .env
├── Automation_Framework.ftsln   UFT solution
├── Driver/               UFT test: Action0 -> Action1 (bootstrap)
├── FunctionLibrary/
│   ├── EnvUtil.vbs       Loads .env, GetConfig(), path helpers
│   ├── ReportUtil.vbs    LogPass/LogFail/RunStep, screenshots, step tracking
│   ├── DataUtil.vbs      Execution CSV reader, TestData parsing, status CSV
│   ├── ORUtil.vbs        Loads shared object repositories (.tsr)
│   ├── BrowserUtil.vbs   OpenBrowser / CloseBrowser driven by .env
│   ├── WebActions.vbs    WebSet, WebSetSecure, WebClick, WebClickRobust, ...
│   ├── Assertions.vbs    VerifyText, VerifyEquals, VerifyUrlContains, ...
│   ├── DriverUtil.vbs    InitFramework, MainDriver, RunTestCase
│   ├── AppFunctions.vbs  YOUR business keywords (Login template)
│   └── Tests.vbs         YOUR test cases (2 samples)
├── ObjectRepository/     Put shared .tsr files here
├── TestData/Execution.csv  Which tests run, and their data
├── Reports/              ExecutionStatus.csv + ExecutionReport.html (generated)
├── Screenshots/          Step screenshots per run (generated)
└── Scripts/
    ├── RunTest.vbs           Command-line runner (returns exit code)
    ├── run_uft.ps1           CI entry point
    └── GenerateHtmlReport.vbs
```

## Quick start

1. Clone the repo anywhere (no fixed path needed).
2. Copy `.env.example` to `.env` and set at least `BASE_URL` (plus `APP_USERNAME` / `APP_PASSWORD` if your app needs a login).
3. Open `Automation_Framework.ftsln` in UFT (or open the `Driver` folder as a test). Make sure the **Web** add-in is loaded and the UFT browser extension is enabled in your browser.
4. Run. `TC01` (`TC_Sample_OpenApplication`) opens `BASE_URL` and checks it loaded, which works against any site.

From the command line:

```
cscript //nologo Scripts\RunTest.vbs
```

## How a run works

`Driver/Action1` finds the framework root (the parent of the `Driver` folder), loads every library in `FunctionLibrary`, then calls:

1. `InitFramework` loads `.env` and the shared object repositories.
2. `MainDriver` reads `TestData/Execution.csv` and, for each row with `Execute_Status = Yes`, opens the browser, calls the Sub named in `Function_Name` with the row's `TestData` as a dictionary, closes the browser and decides **Passed / Failed**.
3. Results go to the UFT run results, `Reports/ExecutionStatus.csv` and (via the runner) `Reports/ExecutionReport.html`.

A test case passes only if at least one step passed, no step failed and no runtime error occurred.

## Configuration (.env)

| Key | Purpose |
|---|---|
| `APP_NAME` | Shown in reports |
| `BASE_URL` | Start URL opened for each test |
| `APP_USERNAME`, `APP_PASSWORD` | Login credentials (passwords are masked in the results) |
| `BROWSER` | `chrome`, `edge` or `firefox` |
| `BROWSER_ARGS` | Extra command-line switches, e.g. `--incognito` |
| `OPEN_BROWSER_PER_TEST` | Open/close the browser around every test case |
| `CLOSE_BROWSERS_BEFORE_TEST`, `CLOSE_ALL_BROWSERS` | Clean-up behaviour |
| `DEFAULT_TIMEOUT`, `PAGE_LOAD_TIMEOUT` | Seconds |
| `EXECUTION_FILE` | Execution CSV (relative to root) |
| `OBJECT_REPOSITORY` | `.tsr` path(s), `;`-separated; empty = all in `ObjectRepository/` |
| `SCREENSHOT_ON_FAIL`, `SCREENSHOT_ON_PASS` | Screenshot policy |
| `RESULTS_DIR`, `UFT_VISIBLE` | Used by `Scripts/RunTest.vbs` |

Rules: `#` starts a comment, values can be quoted, `${OTHER_KEY}` references another key. In code, read values with `GetConfig("KEY")`, `GetConfigInt`, `GetConfigBool` or `RequireConfig` (fails the step if missing).

**Several environments:** set `TEST_ENV=qa` and create `.env.qa` with only the keys that differ; it is loaded on top of `.env`.

**Overrides:** a Windows environment variable with the same name as a key wins over the file. In CI, `Scripts/run_uft.ps1` builds `.env` from `.env.example` plus CI variables when no `.env` exists, so secrets come from the CI credential store and never from git.

## Adding a test

1. Add business keywords to `AppFunctions.vbs` (return `True`/`False`).
2. Add a Sub to `Tests.vbs`:

   ```vb
   Sub TC_CreateOrder(dictData)
       Dim status
       status = Login("", "")                 ' uses .env credentials
       RunStep "Login", status
       If Not status Then Exit Sub
       RunStep "Create order", CreateOrder(dictData("product"), dictData("qty"))
       RunStep "Logout", Logout()
   End Sub
   ```

3. Add a row to `TestData/Execution.csv`:

   ```
   TC03,Yes,TC_CreateOrder,product:Widget;qty:2
   ```

`TestData` is `key:value` pairs separated by `;`. Wrap the field in double quotes if it contains commas. Values may use `${KEY}` from `.env`. A `url` key overrides `BASE_URL` for that test.

To add a new library, put the `.vbs` in `FunctionLibrary/` and add its name to the `libraries` list in `Driver/Action1`.

## CI (Jenkins example)

```groovy
bat 'powershell -ExecutionPolicy Bypass -File Scripts\\run_uft.ps1'
archiveArtifacts artifacts: 'Reports/**, Screenshots/**', allowEmptyArchive: true
```

Exit code `0` means passed; non-zero fails the build. Inject `BASE_URL`, `APP_USERNAME`, `APP_PASSWORD` as environment variables or credentials.

## Pushing to Git

```
git init
git add .
git commit -m "Initial UFT framework"
git remote add origin <your-repo-url>
git push -u origin main
```

`.env`, run results (`Driver/Res*`, `Results/`), reports and screenshots are excluded by `.gitignore`. `.gitattributes` keeps CRLF for scripts and marks UFT files (`.tsp`, `.mtr`, `.bdb`, `.tsr`) as binary so git never tries to merge them.
