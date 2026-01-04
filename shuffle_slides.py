#!/usr/bin/env python3
"""
PowerPoint Slide Shuffler

This script shuffles all slides in a PowerPoint presentation randomly.
"""

import argparse
import random
import sys
from pathlib import Path
from pptx import Presentation


def shuffle_slides(input_file, output_file=None):
    """
    Shuffle all slides in a PowerPoint presentation.
    
    Args:
        input_file (str): Path to the input PowerPoint file
        output_file (str, optional): Path to the output file. If None, creates a new file
                                     with '_shuffled' suffix
    
    Returns:
        str: Path to the output file
    """
    input_path = Path(input_file)
    
    # Validate input file
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")
    
    if input_path.suffix.lower() not in ['.pptx', '.ppt']:
        raise ValueError(f"Input file must be a PowerPoint file (.pptx or .ppt): {input_file}")
    
    # Determine output file path
    if output_file is None:
        output_file = input_path.stem + '_shuffled' + input_path.suffix
    
    print(f"Loading presentation from: {input_file}")
    
    # Load the presentation
    prs = Presentation(input_file)
    
    # Get the number of slides
    num_slides = len(prs.slides)
    print(f"Found {num_slides} slides")
    
    if num_slides < 2:
        print("Warning: Presentation has less than 2 slides. Nothing to shuffle.")
        prs.save(output_file)
        return output_file
    
    # Create a list of slide indices
    slide_indices = list(range(num_slides))
    
    # Shuffle the indices
    random.shuffle(slide_indices)
    
    print(f"Shuffling slides with new order: {slide_indices}")
    
    # Reorder slides by manipulating the slide list directly
    # Note: This uses private API (_sldIdLst) as python-pptx doesn't provide
    # a public method for reordering slides. This is the most efficient approach
    # that avoids duplicating layout resources.
    slides_list = prs.slides._sldIdLst
    
    # Get all slide relationships
    slide_rels = [slides_list[i] for i in range(len(slides_list))]
    
    # Create a new list in shuffled order
    shuffled_rels = [slide_rels[i] for i in slide_indices]
    
    # Clear the current slide list efficiently
    slides_list.clear()
    
    # Add slides back in shuffled order
    for rel in shuffled_rels:
        slides_list.append(rel)
    
    # Save the modified presentation
    print(f"Saving shuffled presentation to: {output_file}")
    prs.save(output_file)
    
    print(f"Successfully created shuffled presentation: {output_file}")
    return output_file


def main():
    """Main entry point for the script."""
    parser = argparse.ArgumentParser(
        description='Shuffle all slides in a PowerPoint presentation',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s presentation.pptx
  %(prog)s presentation.pptx -o shuffled.pptx
  %(prog)s input.pptx --output output.pptx
        """
    )
    
    parser.add_argument(
        'input_file',
        help='Path to the input PowerPoint file (.pptx or .ppt)'
    )
    
    parser.add_argument(
        '-o', '--output',
        dest='output_file',
        help='Path to the output file (default: <input>_shuffled.pptx)'
    )
    
    parser.add_argument(
        '--seed',
        type=int,
        help='Random seed for reproducible shuffling'
    )
    
    args = parser.parse_args()
    
    # Set random seed if provided
    if args.seed is not None:
        random.seed(args.seed)
        print(f"Using random seed: {args.seed}")
    
    try:
        output_file = shuffle_slides(args.input_file, args.output_file)
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
