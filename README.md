# AI-Powered PDF to Excel Extraction Tool

This repository contains an advanced AI-powered tool for extracting tables from PDF files and converting them to Excel format (.xlsx). The tool uses machine learning algorithms to intelligently detect and extract tabular data with high accuracy.

## 🚀 Features

- **AI-Powered Table Detection**: Uses advanced machine learning algorithms (via Camelot) for intelligent table recognition
- **Multiple Parsing Strategies**: Automatically tries different extraction methods for optimal results
- **Excel Output**: Exports to native Excel format (.xlsx) with proper formatting
- **Intelligent Data Cleaning**: AI-enhanced data preprocessing and validation
- **Batch Processing**: Process multiple PDF files simultaneously
- **Error Handling**: Comprehensive error handling and logging
- **Progress Tracking**: Real-time progress reporting and detailed logs

## 📦 Quick Start

### Option 1: Automated Setup
```bash
chmod +x setup.sh
./setup.sh
```

### Option 2: Manual Setup
```bash
# Install dependencies
pip3 install camelot-py pandas openpyxl tabulate PyPDF2 numpy

# Create directories
mkdir -p input_pdfs output_excel
```

## 🎯 Usage

### Basic Usage
1. Place your PDF files in the `input_pdfs` directory
2. Run the extraction tool:
   ```bash
   python3 pdf_to_excel_ai.py
   ```
3. Check the `output_excel` directory for your Excel files

### Advanced Usage
```bash
# Custom input/output directories
python3 pdf_to_excel_ai.py -i /path/to/pdfs -o /path/to/output

# Verbose logging for debugging
python3 pdf_to_excel_ai.py --verbose

# Get help
python3 pdf_to_excel_ai.py --help
```

### Demo Mode
Run a quick demonstration:
```bash
python3 demo.py
```

## 🧠 AI Features

The tool incorporates several AI-powered features:

1. **Smart Table Detection**: Uses multiple machine learning models to detect tables with different structures
2. **Intelligent Parsing**: Automatically selects the best parsing strategy based on document analysis
3. **Data Quality Enhancement**: AI-powered data cleaning and validation
4. **Format Recognition**: Automatically detects headers, data types, and table structure
5. **Error Recovery**: Intelligent fallback strategies when primary methods fail

## 📊 Output Format

- **Excel Files**: Native .xlsx format with proper formatting
- **Multiple Sheets**: Each table becomes a separate worksheet
- **Auto-sizing**: Column widths automatically adjusted for readability
- **Data Preservation**: Maintains original data integrity while cleaning noise

## 🔧 Technical Details

### Dependencies
- `camelot-py`: Core PDF table extraction engine
- `pandas`: Data manipulation and analysis
- `openpyxl`: Excel file creation and formatting
- `PyPDF2`: PDF validation and metadata extraction
- `numpy`: Numerical operations and data cleaning

### Parsing Strategies
1. **Stream-based**: For tables without clear borders
2. **Lattice-based**: For tables with defined grid lines
3. **Custom Areas**: For complex layouts with manual area definition

### AI Algorithms
- Machine learning-based table boundary detection
- Intelligent text classification and grouping
- Automated data type inference
- Smart header detection and separation

## 📁 File Structure

```
├── pdf_to_excel_ai.py      # Main extraction script
├── demo.py                 # Demonstration script
├── setup.sh                # Automated setup script
├── requirements.txt        # Python dependencies
├── input_pdfs/             # Place PDF files here
├── output_excel/           # Excel files generated here
├── examples/               # Example notebooks and samples
└── README.md              # This file
```

## 🚀 Examples

Check the `examples/` directory for:
- Jupyter notebook with advanced usage patterns
- Sample PDF files for testing
- Advanced configuration examples

---

# Configuring a Repl

Every new repl comes with a `.replit` and a `replit.nix` file that let you configure your repl to do just about anything in any language!

### `replit.nix`

Every new repl is now a Nix repl, which means you can install any package available on Nix, and support any number of languages in a single repl. You can search for a list of available packages [here](https://search.nixos.org/packages).

The `replit.nix` file should look something like the example below. The `deps` array specifies which Nix packages you would like to be available in your environment. 

```nix
{ pkgs }: {
    deps = [
        pkgs.python312
        pkgs.python312Packages.pip
        pkgs.python312Packages.pandas
        pkgs.python312Packages.numpy
        pkgs.python312Packages.openpyxl
        pkgs.python312Packages.requests
        pkgs.ghostscript
        pkgs.tkinter
    ];
}
```
### Learn More About Nix

If you'd like to learn more about Nix, here are some great resources:

#### Written Guides
- [Getting started with Nix](/programming-ide/getting-started-nix) — Our own getting started guide
- [Building with Nix on Replit](https://docs.replit.com/tutorials/30-build-with-nix) — Deploy a production web stack on Replit with Nix
- [Nix Pills](https://nixos.org/guides/nix-pills/) — Guided introduction to Nix
- [Nix Package Manager Guide](https://nixos.org/manual/nix/stable/) — A comprehensive guide of the Nix Package Manager
- [A tour of Nix](https://nixcloud.io/tour) — Learn the nix language itself

#### Video Guides
- [Nixology](https://www.youtube.com/playlist?list=PLRGI9KQ3_HP_OFRG6R-p4iFgMSK1t5BHs) — A series of videos introducing Nix in a practical way
- [Taking the Nix pill](https://www.youtube.com/watch?v=QwLWIy2KleE) — An introduction to what Nix is, how it works, and a walkthrough of publishing several new languages to Replit within an hour.
- [Nix: A Deep Dive](https://www.youtube.com/watch?v=TsZte_9GfPE) — A deep dive on Nix: what Nix is, why you should use it, and how it works.


### `.replit`

The `.replit` file allows you to configure many options for your repl, most basic of which is the `run` command.

Check out how to use the `.replit` file to configure a repl to enable [Clojure](https://clojure.org):

<iframe width="640" height="400" style="margin-bottom: 10px;" src="https://www.loom.com/embed/cbe1f74399c546c38e0c1871893816c5" frameborder="0" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>

`.replit` files follow the [toml configuration format](https://learnxinyminutes.com/docs/toml/) and look something like this:

```toml
# The command that is executed when the run button is clicked.
run = ["cargo", "run"]

# The default file opened in the editor.
entrypoint = "src/main.rs"

# Setting environment variables
[env]
FOO="foo"

# Packager configuration for the Universal Package Manager
# See https://github.com/replit/upm for supported languages.
[packager]
language = "rust"

  [packager.features]
  # Enables the package search sidebar
  packageSearch = true
  # Enabled package guessing
  guessImports = false

# Per language configuration: language.<lang name> 
[languages.rust]
# The glob pattern to match files for this programming language
pattern = "**/*.rs"

    # LSP configuration for code intelligence
    [languages.rust.languageServer]
    start = ["rust-analyzer"]
```

In the code above, the strings in the array assigned to `run` are executed in order in the shell whenever you hit the "Run" button. 

The `language` configuration option helps the IDE understand how to provide features like [packaging](https://blog.replit.com/upm) and [code intelligence](https://blog.replit.com/intel).

And the `[languages.rust]` `pattern` option is configured so that all files ending with `.rs` are treated as Rust files. The name is user-defined and doesn't have any special meaning, we could have used `[languages.rs]` instead.

We can now set up a language server specifically for Rust. Which is what we do with the next configuration option: `[languages.rust.languageServer]`. [Language servers](https://microsoft.github.io/language-server-protocol/#:~:text=A%20Language%20Server%20is%20meant,servers%20and%20development%20tools%20communicate.) add smart features to your editor like code intelligence, go-to-definition, and documentation on hover.

Since repls are fully configurable, you're not limited to just one language. For example, you could install Clojure and its language server using `replit.nix`, add a `[languages.clojure]` configuration option to the above `.replit` file that matched all Clojure files and have code intelligence enabled for both languages in the same repl.

### `.replit` reference

A `Command` can either be a string or a list of strings. If the `Command` is a string (`"node index.js"`), it will be executed via `sh -c "<Command>"`. If the Command is a list of strings (`["node", "index.js"]`), it will be directly executed with the list of strings passed as arguments. When possible, it is preferred to pass a list of strings.

- `run`
  - **Type:** `Command`
  - **Description:** The command to run the repl.
- `entrypoint`
  - **Type:**  `string`
  - **Description:** The name of the main file including the extension. This is the file that will be run, and shown by default when opening the editor.
- `onBoot`
  - **Type:** `Command`
  - **Description:** The command that executes after your repl has booted.
- `compile`
  - **Type:** `Command`
  - **Description:** The shell command to compile the repl before the `run` command. Only for compiled languages like C, C++, and Java.
- `audio`
  - **Type:** `boolean`
  - **Description:** Enables [system-wide audio](https://docs.replit.com/misc/playing-audio-replit) for the repl when configured to `true`.
- `language`
  - **Type:** `string`
  - **Description:** Reserved. During a GitHub import, this tells the workspace which language should be used when creating the repl. For new repls, this option will always be Nix, so this field should generally not be touched.
- `[env]`
  - **Description:** Set environment variables. Don't put secrets here—use the Secrets tab in the left sidebar.
  - **Example:** `VIRTUAL_ENV = "/home/runner/${REPL_SLUG}/venv"`
- `interpreter`
  - **Description:** Specifies the interpreter, which should be a compliant [prybar binary](https://github.com/replit/prybar).
  - `command`
    - **Type:** `[string]`
    - **Description:** This is the command that will be run to start the interpreter. It has higher precedence than the `run` command (i.e. `interpreter` command will run instead of the `run` command).
  - `prompt`
    - **Type:** `[byte]`
    - **Description:** This is the prompt used to detect running state, if unspecified it defaults to `[0xEE, 0xA7]`.
- `[unitTest]`
  - Enables unit testing to the repl.
  - `language`
      - **Type:** `string`
      - **Description:** The language you want the unit tests to run. Supported strings: `java`, `python`, and `nodejs`.
- `[packager]`
  - **Description:** Package management configuration. Learn more about installing packages [here](https://docs.replit.com/repls/packages/#DirectImports).
  - `afterInstall`
    - **Type:** `Command`
    - **Description:** The command that is executed after a new package is installed.
  - `ignoredPaths`
    - **Type:** `[string]`
    - **Description:** List of paths to ignore while attempting to guess packages.
  - `ignoredPackages`
    - **Type:** `[string]`
    - **Description:** List of modules to never attempt to guess a package for, when installing packages.
  - `language`
    - **Type:** `string`
    - **Description:** Specifies the language to use for package operations. See available languages in the [Universal Package Manager](https://github.com/replit/upm) repository.
  - `[packager.features]`
    - **Description:** UPM features that are supported by the specified languages.
      - `packageSearch`
        - **Type:** Boolean
        - **Description:** When set to `true`, enables a package search panel in the sidebar.
      - `guessImports`
        - **Type:** Boolean
        - **Description:** When set to `true`, UPM will attempt to guess which packages need to be installed prior to running the repl.
- `[languages.<language name>]`
  - **Description:** Per-language configuration. The language name has no special meaning other than to allow multiple languages to be configured at once.
  - `pattern`
    - **Type:** `string`
    - **Description:** A [glob](https://en.wikipedia.org/wiki/Glob_(programming)) used to identify which files belong to this language configuration (`**/*.js`)
  - `syntax`
    - **Type:** `string`
    - **Description:** The language to use for syntax highlighting.
  - `[languages.<language name>.languageServer]`
    - **Description:** Configuration for setting up [LSP](https://microsoft.github.io/language-server-protocol/) for this language. This allows for code intelligence (autocomplete, underlined errors, etc...).
    - `start`
      - **Type:** `Command`
      - **Description:** The command used to start the LSP server for the specified language.
- `[nix]`
  - **Description:** Where you specify a [Nix channel](https://nixos.wiki/wiki/Nix_channels).
  - `channel`
    - **Type:** `string`
    - **Description:** A nix channel id.
- `[debugger]`
  - **Description:** Advanced users only. See field types & docstrings [here](https://gist.github.com/Bardia95/98987c69c6970b1bb0698b863e2a84de#file-dot-replit-debugger-config-go), and in the advanced section below.

### Example configurations
#### Beginner
##### [LaTeX](https://replit.com/@ZachAtReplit/LaTeX?v=1#.replit)
##### [Clojure](https://replit.com/@replit/Clojure?v=1#.replit)
#### Advanced
##### [Python](https://replit.com/@replit/Python?v=1)
##### [HTML, CSS, JS](https://replit.com/@replit/HTML-CSS-JS?v=1#.replit)
##### [Java](https://replit.com/@replit/Java-Beta?v=1#.replit)
##### [Node.js](https://replit.com/@replit/Nodejs?v=1#.replit)
##### [C++](https://replit.com/@replit/CPlusPlus?v=1)