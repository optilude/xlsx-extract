import os
import sys
import argparse
import zipfile

import openpyxl
import openpyxl.utils.exceptions

from .config import run

def build_arg_parser():

    parser = argparse.ArgumentParser()

    parser.add_argument('target', metavar="target.xlsx", help="The Excel workbook defining the format the data should be extracted to.")
    parser.add_argument('output', nargs="?", metavar="output.xlsx", help="Output file where results will be written. Required unless --update is given.")
    parser.add_argument('--update', action='store_true', help="Use this instead of naming an output file to overwrite the target with the extract results.")
    parser.add_argument('--allow-failures', action='store_true', help="By default, the output file will not be written if there are any failures in the extract. Set this option to write even if some extracts failed.")
    parser.add_argument('--config-sheet', metavar="Config", default="Config", help="Name of the worksheet in the target file that contains the extract configuration.")
    parser.add_argument('--source-directory', metavar="/path/to/source", help="Directory where source files are found. Defaults to the current directory, and can be overridden in the extract configuration.")
    parser.add_argument('--source-file', metavar="source.xlsx", help="Source file to extract from. Can be overridden in the extract configuration.")

    return parser

def main():
    parser = build_arg_parser()
    args = parser.parse_args()

    cwd = os.getcwd()
    
    target_filename = args.target
    output_filename = args.output if not args.update else args.target
    source_directory = args.source_directory if args.source_directory else cwd
    source_file = args.source_file
    config_sheet = args.config_sheet
    allow_failures = args.allow_failures

    if not target_filename or not output_filename or (args.update and args.output):
        parser.print_usage()
        return
    
    if not os.path.isfile(target_filename):
        print("Target file %s not found" % target_filename)
        return
    
    if not os.path.isdir(source_directory):
        print("Source directory %s not found" % source_directory)
        return
        
    if source_file:
        if not os.path.isfile(source_file):
            print("Source file %s not found" % source_file)
            return

    target_workbook = None

    try:
        target_workbook = openpyxl.load_workbook(args.target, data_only=False)
    except (openpyxl.utils.exceptions.InvalidFileException, zipfile.BadZipFile, FileNotFoundError,) as e:
        print(str(e))
        return

    # operate from the source directory, if given
    os.chdir(source_directory)
    
    history = run(target_workbook, source_directory, source_file, config_sheet)

    success = True
    for action in history:
        print(action)

        if not action.success:
            success = False

    # Change back to the current working directory for writing output
    os.chdir(cwd)

    if success or allow_failures:
        target_workbook.save(output_filename)

    sys.exit(0 if success else 1)

if __name__ == '__main__':
    main()
