"""Translate PPTX files from English to Chinese.

This script:
1. Extracts text from a PPTX file
2. Translates the text to Chinese
3. Creates a new PPTX with translated text

Usage:
    python translate_pptx.py input.pptx -o output.pptx
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Dict, List, Any
from copy import deepcopy

from pptx import Presentation
from translate import Translator

def translate_text(text: str) -> str:
    """Translate text from English to Chinese using translate library."""
    translator = Translator(to_lang="zh")
    try:
        return translator.translate(text)
    except Exception as e:
        print(f"Warning: Failed to translate text: {text}")
        print(f"Error: {e}")
        return text

def translate_presentation(input_path: Path, output_path: Path) -> None:
    # Load the presentation
    prs = Presentation(str(input_path))
    
    # Process each slide
    for slide in prs.slides:
        # Translate title if exists
        if slide.shapes.title is not None:
            title_text = slide.shapes.title.text
            if title_text.strip():
                slide.shapes.title.text = translate_text(title_text)
        
        # Process each shape in the slide
        for shape in slide.shapes:
            # Skip the title shape since we handled it above
            if shape == slide.shapes.title:
                continue
                
            # Translate text frames
            if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
                text = shape.text_frame.text
                if text.strip():
                    # Clear existing paragraphs
                    while len(shape.text_frame.paragraphs) > 1:
                        p = shape.text_frame.paragraphs[-1]
                        p._p.getparent().remove(p._p)
                    
                    # Set translated text
                    shape.text_frame.text = translate_text(text)
            
            # Translate tables
            if hasattr(shape, 'has_table') and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            cell.text = translate_text(cell.text)
        
        # Translate notes if they exist
        if slide.has_notes_slide:
            notes_slide = slide.notes_slide
            if notes_slide and notes_slide.notes_text_frame:
                notes_text = notes_slide.notes_text_frame.text
                if notes_text.strip():
                    notes_slide.notes_text_frame.text = translate_text(notes_text)
    
    # Save the translated presentation
    prs.save(str(output_path))

def cli_main() -> None:
    parser = argparse.ArgumentParser(description="Translate PPTX from English to Chinese")
    parser.add_argument("input", help="Path to input .pptx file")
    parser.add_argument("-o", "--output", help="Path to output .pptx file")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        raise SystemExit(f"Input file not found: {input_path}")

    if args.output:
        output_path = Path(args.output)
    else:
        # Create default output filename by adding _zh suffix
        output_path = input_path.parent / f"{input_path.stem}_zh{input_path.suffix}"

    translate_presentation(input_path, output_path)
    print(f"Created translated presentation: {output_path}")

if __name__ == "__main__":
    cli_main()