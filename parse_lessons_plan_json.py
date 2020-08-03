
from lessons_parser.lessons_parser import *
import argparse

parser = argparse.ArgumentParser(description='Python program for upload bysy time for rooms')
parser.add_argument('-in','--input', metavar='input',help='Input .xlsx file for uploading tables', required=True)
parser.add_argument('-out','--output', metavar='output',help='Output .json file for uploading', required=True)
parser.add_argument('-log','--outputlog', metavar='log output',help='Output log file for uploading')
parser.add_argument('-tc','--time_config', metavar='time_config',help='Input file for times info', required=True)


args = parser.parse_args()
input_filename=args.input;
output_filename=args.output;
time_config=args.time_config
log_filename=args.outputlog;

lp = LessonsParser(input_filename, time_config)
lp.parse()
lp.save(output_filename)
