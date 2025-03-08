# Flowchart to PowerPoint Converter

This tool converts flowchart images into editable PowerPoint presentations. It detects shapes, extracts text using OCR, and preserves connections between shapes.

## Features

- Detects common flowchart shapes (rectangles, diamonds, ovals)
- Extracts text from shapes using OCR
- Preserves connections between shapes
- Creates fully editable PowerPoint slides
- Maintains relative positions and sizes of elements

## Prerequisites

- Python 3.7 or higher
- Tesseract OCR engine installed on your system

### Installing Tesseract OCR

#### Windows
1. Download the installer from [GitHub](https://github.com/UB-Mannheim/tesseract/wiki)
2. Run the installer
3. Add Tesseract to your system PATH

#### Linux
```bash
sudo apt-get install tesseract-ocr
```

#### macOS
```bash
brew install tesseract
```

## Installation

1. Clone this repository:
```bash
git clone [repository-url]
cd [repository-name]
```

2. Install required Python packages:
```bash
pip install -r requirements.txt
```

## Usage

Run the script from the command line:

```bash
python flowchart_to_pptx.py <input_image> <output_pptx>
```

Example:
```bash
python flowchart_to_pptx.py flowchart.png output.pptx
```

## Input Requirements

- Supported image formats: PNG, JPEG, BMP
- Clear, high-contrast flowchart images work best
- Shapes should be clearly defined
- Text should be readable and not overlapping

## Output

The tool generates a PowerPoint file (.pptx) where:
- Each shape is independently editable
- Text can be modified
- Connectors can be adjusted
- Shapes can be moved while maintaining connections

## Limitations

- Complex or hand-drawn flowcharts may not be detected accurately
- Very small text might not be recognized correctly
- Custom or unusual shapes might be simplified to basic shapes
- Overlapping shapes may cause detection issues

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. 