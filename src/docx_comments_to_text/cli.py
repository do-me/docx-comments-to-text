import click
import sys
from pathlib import Path
from .docx_processor import process_docx

@click.command()
@click.argument('input_file', type=click.Path(exists=True, path_type=Path))
@click.option('-o', '--output', 'output_file', type=click.Path(path_type=Path), 
              help='Output file path. If not specified, prints to stdout.')
@click.option('--authors', type=click.Choice(['never', 'always', 'auto']), default='always',
              help='How to display comment authors (default: always)')
@click.option('--placement', type=click.Choice(['inline', 'end-paragraph', 'comments-only']), default='inline',
              help='Comment placement style (default: inline)')
def main(input_file: Path, output_file: Path, authors: str, placement: str):
    """Extract comments from DOCX files and insert them inline with the text."""
    
    try:
        # Process the DOCX file
        formatted_text = process_docx(input_file, authors, placement)
        
        # Output the result
        if output_file:
            output_file.write_text(formatted_text, encoding='utf-8')
            click.echo(f"Output written to: {output_file}")
        else:
            click.echo(formatted_text)
            
    except FileNotFoundError as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)
    except Exception as e:
        click.echo(f"Error processing file: {e}", err=True)
        sys.exit(1)

if __name__ == "__main__":
    main()
