#!/usr/bin/env python
# -*- coding: utf-8 -*-

from zipfile import ZipFile
import argparse
import os

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Peter Mayor"


def find_all(a_str, sub):
    start = 0
    while True:
        start = a_str.find(sub, start)
        if start == -1:
            return
        yield start
        start += len(sub)  # use start += 1 to find overlapping matches


def search_zipped_file(file_name, search_text, files_with_string_list):
    with ZipFile(file_name, 'r') as zip:
        print(file_name)
        with zip.open("word/document.xml") as contents:
            content_string = contents.read()
            string_match_count = 0
            for found_index in find_all(content_string, search_text):
                string_match_count += 1
                files_with_string_list.append(file_name)
                print(content_string[found_index-40:found_index+40])


def create_parser():
    parser = argparse.ArgumentParser(
        description="searches for text in dotm files in a folder")
    parser.add_argument(
        "--dir", help="takes in directory to search", default=".")
    parser.add_argument("text", help="text to search for")
    return parser


def directory_loop(args):
    file_list = os.listdir(args.dir)
    files_with_string_list = []
    searched_file_count = 0
    for file in file_list:
        full_path = os.path.join(args.dir, file)
        if full_path.endswith('.dotm'):
            search_zipped_file(full_path, args.text, files_with_string_list)
            searched_file_count += 1
    files_with_string_set = set(files_with_string_list)
    print("\n \n")
    print("Files with Matches: \n" + "\n".join(list(files_with_string_set)))
    print("Number of files with Matches: " + str(len(files_with_string_set)))
    print("Files Searched: " + str(searched_file_count))


def main():
    parser = create_parser()
    args = parser.parse_args()
    directory_loop(args)


if __name__ == '__main__':
    main()
