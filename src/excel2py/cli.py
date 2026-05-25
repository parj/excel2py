from __future__ import annotations

import logging
import sys
from pathlib import Path

import click

from excel2py.converter import convert


@click.group()
@click.version_option(package_name="excel2py")
def main():
    """excel2py - Convert Excel spreadsheets to Python scripts using GenAI."""


@main.command()
@click.argument("input_file", type=click.Path(exists=True, path_type=Path))
@click.option(
    "-o",
    "--output",
    "output_file",
    type=click.Path(path_type=Path),
    default=None,
    help="Output Python file path. Defaults to <input>_converted.py",
)
@click.option(
    "-p",
    "--provider",
    type=click.Choice(["openai", "anthropic", "google", "openrouter"]),
    default=None,
    help="LLM provider to use",
)
@click.option("-m", "--model", default=None, help="Model override")
@click.option("--api-key", default=None, help="API key (overrides env config)")
@click.option("--dry-run", is_flag=True, help="Parse Excel and print prompt without calling LLM")
@click.option(
    "--no-verify", "no_verify", is_flag=True, help="Skip verification-and-correction loop"
)
@click.option(
    "--max-verify-attempts", type=int, default=None, help="Max correction attempts (default 3)"
)
@click.option("-v", "--verbose", is_flag=True, help="Enable debug logging")
def convert_cmd(
    input_file,
    output_file,
    provider,
    model,
    api_key,
    dry_run,
    no_verify,
    max_verify_attempts,
    verbose,
):
    """Convert an Excel spreadsheet to a Python script."""
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    if output_file is None:
        output_file = input_file.with_name(f"{input_file.stem}_converted.py")

    from excel2py.config import get_settings

    settings = get_settings()

    try:
        result = convert(
            input_file=input_file,
            output_file=None if dry_run else output_file,
            provider=provider,
            model=model,
            api_key=api_key,
            dry_run=dry_run,
            verify=not no_verify,
            max_verify_attempts=max_verify_attempts
            if max_verify_attempts is not None
            else settings.max_verify_attempts,
            verify_timeout=settings.verify_timeout,
            settings=settings,
        )
        if dry_run:
            click.echo(result)
        else:
            click.echo(f"Converted {input_file} -> {output_file}")
    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)
