# Make Video Previews

Make Video Previews turns a folder full of video files into a Word document (`.docx`) with many preview thumbnails per page. It is meant for people who need a quick visual overview of footage without opening every clip one by one.

![Preview document example](./assets/readme-preview-nasa-webb-star-factory.png)

Preview frames from [NASA Image and Video Library: "The Webb Space Telescope’s New Look at a “Star Factory” on This Week @NASA – October 21, 2022"](https://images.nasa.gov/details/The%20Webb%20Space%20Telescope%E2%80%99s%20New%20Look%20at%20a%20%E2%80%9CStar%20Factory%E2%80%9D%20on%20This%20Week%20%40NASA%20%E2%80%93%20October%2021%2C%202022), exported with makevideopreviews. NASA is the source of the imagery.

## What It Does

You point the tool at a folder that contains your project, card dumps, archive material, or other video folders.

It then:

- searches that folder and all subfolders for video files
- extracts preview frames from the videos at a fixed time interval
- places many thumbnails on landscape Word pages
- writes the clip name and timecode directly onto each thumbnail
- creates one preview document for the whole project, or one preview document per video subfolder

## What The Result Looks Like

The output is a Word file with:

- landscape pages
- many small thumbnails per page
- very little wasted space
- one dark text overlay on each thumbnail
- a full timecode on every thumbnail
- shortened clip names if many files share the same prefix

This is especially useful for documentary footage, film material, archive material, and larger shooting days with many different visual situations. For interviews it can still help as an overview, but the strongest use case is footage where the image changes a lot over time.

## Who This Is For

This tool is useful if you:

- work with large amounts of footage
- need a quick overview before editing or logging
- want to send a visual shot list to someone else
- prefer a simple Word document over a special media tool

## What You Need

You need:

- Python 3
- `ffmpeg`
- `ffprobe`
- the Python packages used by this project

On macOS, `ffmpeg` is often installed with Homebrew:

```bash
brew install ffmpeg
```

## Installation

Open Terminal, go into the project folder, and run:

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install --upgrade pip
python3 -m pip install -e .
```

After that, the command `makevideopreviews` is available inside that environment.

## First Check

To check whether everything required is installed, run:

```bash
makevideopreviews doctor
```

The doctor command checks:

- `ffmpeg`
- `ffprobe`
- required Python packages
- write access

## Easiest Way To Use It

Start the interactive mode:

```bash
makevideopreviews
```

Then answer a few simple questions:

- which folder should be scanned
- whether you want one project document or separate subfolder documents
- how many seconds should lie between preview frames
- which image quality you want
- whether existing preview files should be overwritten

## Fast Command-Line Examples

Create one preview file for a whole project:

```bash
makevideopreviews generate --root "/path/to/project"
```

Only estimate the size first:

```bash
makevideopreviews estimate --root "/path/to/project"
```

Create one preview document per subfolder instead:

```bash
makevideopreviews generate --root "/path/to/project" --scope subfolder
```

Overwrite an existing preview file:

```bash
makevideopreviews generate --root "/path/to/project" --overwrite
```

## Main Options

### `--root`

The main folder that contains your project or footage tree.

### `--scope`

Controls how many Word files are created.

- `project`: one Word file for the whole project
- `subfolder`: one Word file per video subfolder

Default: `project`

### `--interval`

How often the tool takes a preview frame.

Example:

- `10` means one preview every 10 seconds

Smaller numbers:

- more thumbnails
- more detail
- bigger output files
- slower runs

Larger numbers:

- fewer thumbnails
- smaller files
- faster runs

### `--quality`

How strongly the preview images are compressed.

If you use the interactive mode, you can choose from friendly presets.

### `--workers`

Optional manual override for parallel processing.

In normal use, the tool chooses this automatically.

### `--overwrite`

Replace an existing preview document.

### `--skip-existing`

Quietly leave existing preview documents untouched.

## What Happens During A Run

Before creating the document, the tool shows a preflight summary.

This summary includes:

- how many folders were found
- how many videos were found
- the total duration
- the estimated number of thumbnails
- the estimated output size

Then the preview document is created.

## Output Location

### Project mode

If you choose project mode, the output file is written into the folder you selected with `--root`.

Example:

```text
/MyProject/preview_MyProject.docx
```

### Subfolder mode

If you choose subfolder mode, each output file is written into the corresponding video folder.

## Development

Run the test suite:

```bash
python3 -m unittest discover -s tests -v
```
