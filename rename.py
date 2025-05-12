
from openpyxl import load_workbook
import argparse


def create_cmd_line_parser():
    parser = argparse.ArgumentParser(description='Rename serial letter names ')
    parser.add_argument('filename', metavar='filename')
    parser.add_argument('directory', metavar='directory')

    return parser

def open

if __name__ == "__main__":

    p = create_cmd_line_parser()
    args = p.parse_args()