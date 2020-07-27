from rooms_parser.rooms_parser import *
import argparse

parser = argparse.ArgumentParser(description='Python program for upload bysy time for rooms')
parser.add_argument('-in','--input', metavar='input',help='Input .docx file for uploading tables', required=True)
parser.add_argument('-out','--output', metavar='output',help='Output .json file for uploading', required=True)
parser.add_argument('-rinfo','--rooms_info', metavar='rooms_info',help='Input file for rooms info', required=True)
parser.add_argument('-tc','--time_config', metavar='time_config',help='Input file for times info', required=True)
parser.add_argument('-log','--outputlog', metavar='log output',help='Output log file for uploading')

def main():
    args = parser.parse_args()
    input_filename=args.input
    output_filename=args.output
    log_filename=args.outputlog
    rooms_info = args.rooms_info
    time_config = args.time_config

    rp = RoomsParser(input_filename, rooms_info, time_config)
    rp.parseRoomsTable()
    # rp.parse()
    rp.save(output_filename)


if __name__ == '__main__':
    main()
