from table_parser.table_parser import *
import argparse

parser = argparse.ArgumentParser(description='Python program for upload and collect data from table of study programs plan')
parser.add_argument('-in','--input', metavar='input',help='Input .xlsx file for uploading', required=True)
parser.add_argument('-out','--output', metavar='output',help='Output .json file for uploading', required=True)


def main():
    args = parser.parse_args()
    input_filename=args.input;
    output_filename=args.output;

    tpp = TableParserPlan(input_filename,22)
    tpp.parse()
    tpp.save(output_filename)


if __name__ == '__main__':
    main()
