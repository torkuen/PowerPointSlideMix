# PowerPointSlideMix

A simple Python script to shuffle all slides in a PowerPoint presentation.

## Installation

Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Shuffle slides in a PowerPoint file (creates a new file with `_shuffled` suffix):

```bash
python shuffle_slides.py presentation.pptx
```

### Specify Output File

Shuffle slides and save to a specific output file:

```bash
python shuffle_slides.py presentation.pptx -o output.pptx
```

### Reproducible Shuffling

Use a random seed for reproducible results:

```bash
python shuffle_slides.py presentation.pptx --seed 42
```

## Command-Line Options

- `input_file`: Path to the input PowerPoint file (.pptx)
- `-o, --output`: Path to the output file (optional, default: `<input>_shuffled.pptx`)
- `--seed`: Random seed for reproducible shuffling (optional)

## Examples

```bash
# Shuffle presentation.pptx and create presentation_shuffled.pptx
python shuffle_slides.py presentation.pptx

# Shuffle and save to a specific file
python shuffle_slides.py slides.pptx -o shuffled_slides.pptx

# Reproducible shuffle with seed
python shuffle_slides.py presentation.pptx --seed 12345
```

## Requirements

- Python 3.6+
- python-pptx

## License

MIT