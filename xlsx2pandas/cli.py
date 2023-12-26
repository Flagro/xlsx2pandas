import argparse


def get_args():
    parser = argparse.ArgumentParser(description='Extract pandas dataframes from excel files')
    parser.add_argument('path', help='Path to the file or directory')
    parser.add_argument('--sheets', help='Comma-separated list of sheet names', default=None)
    return parser.parse_args()
