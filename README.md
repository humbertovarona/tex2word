# tex2word

LaTeX Equation to Word Document

This script generates a Microsoft Word document (`.docx`) containing a LaTeX equation in both text and image format. The image is created using Matplotlib and is embedded in the Word document.

## Requirements

Make sure you have the following Python packages installed:

- `matplotlib`
- `python-docx`

You can install these packages using pip:

```bash
pip install matplotlib python-docx
```

## Usage

To run the script, execute it from the command line with a LaTeX equation as the argument. The LaTeX equation should be provided as a single string.

### Example:

```bash
python script_name.py "E = mc^2"
```

Replace `script_name.py` with the actual name of the script file.

## Output

The script will create a Word document named `equation_output.docx`, containing:

1. The LaTeX equation in text format.
2. An image of the rendered LaTeX equation.

## How It Works

1. The script takes a LaTeX equation as a command-line argument.
2. It generates a Word document with the equation as plain text.
3. It uses Matplotlib to render the equation as an image.
4. The image is then embedded into the Word document.
5. If there's an error rendering the LaTeX, an error message is added to the document.

## Notes

- The script uses the `Agg` backend for Matplotlib to render the image in environments without a display server.
- Make sure that LaTeX syntax is correctly formatted for Matplotlib's math rendering capabilities.

## Troubleshooting

If you encounter issues rendering the LaTeX equation, ensure that:

1. You have a LaTeX-compatible equation.
2. The required packages (`matplotlib` and `python-docx`) are installed correctly.

## License

This project is licensed under the MIT License - see the CC0 LICENSE file for details.

