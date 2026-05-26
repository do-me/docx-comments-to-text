from click.testing import CliRunner
from pathlib import Path
import tempfile
import os
from docx_comments_to_text.cli import main


class TestCLI:
    def setup_method(self):
        self.runner = CliRunner()
        self.test_docx = Path(__file__).parent / "docs" / "simple_comment.docx"

    def test_cli_stdout_output(self):
        """CLI prints markdown with XML-tagged comments to stdout by default."""
        result = self.runner.invoke(main, [str(self.test_docx)])

        assert result.exit_code == 0
        assert "<comment" in result.output
        assert "</comment>" in result.output
        assert "[world]" in result.output

    def test_cli_file_output(self):
        """CLI writes markdown with XML-tagged comments when -o is supplied."""
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.md') as tmp:
            output_path = tmp.name

        try:
            result = self.runner.invoke(main, [str(self.test_docx), '-o', output_path])

            assert result.exit_code == 0
            assert f"Output written to: {output_path}" in result.output

            with open(output_path, 'r', encoding='utf-8') as f:
                content = f.read()
            assert "<comment" in content
            assert "[world]" in content
        finally:
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_cli_authors_never(self):
        """--authors never produces tags without an author attribute."""
        result = self.runner.invoke(main, [str(self.test_docx), '--authors', 'never'])

        assert result.exit_code == 0
        assert "<comment>" in result.output
        assert "author=" not in result.output

    def test_cli_authors_always_default(self):
        """The default --authors mode is 'always', so author= appears in the tag."""
        result = self.runner.invoke(main, [str(self.test_docx)])

        assert result.exit_code == 0
        assert 'author="' in result.output

    def test_cli_nonexistent_file(self):
        result = self.runner.invoke(main, ['nonexistent.docx'])

        assert result.exit_code != 0
        assert "Error:" in result.output

    def test_cli_help(self):
        result = self.runner.invoke(main, ['--help'])

        assert result.exit_code == 0
        assert "Extract comments from DOCX files" in result.output
        assert "--authors" in result.output
        assert "--output" in result.output
